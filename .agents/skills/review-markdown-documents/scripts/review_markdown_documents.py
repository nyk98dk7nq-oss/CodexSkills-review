#!/usr/bin/env python3
"""Prepare, validate, and summarize deterministic design-review data."""

from __future__ import annotations

import argparse
import json
import os
import re
import sys
import uuid
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any


RESULT_VALUES = ("適合", "不適合", "対象外", "要確認")
ITEM_KEY_FIELDS = ("checklist_file", "sheet", "row", "target_file")


class ReviewDataError(ValueError):
    """Raised when a manifest or review result is invalid."""


def _is_link_like(path: Path) -> bool:
    return path.is_symlink() or path.is_junction()


def _read_json(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except FileNotFoundError as exc:
        raise ReviewDataError(f"ファイルが見つかりません: {path}") from exc
    except json.JSONDecodeError as exc:
        raise ReviewDataError(
            f"JSONを解析できません: {path} ({exc.lineno}行{exc.colno}列)"
        ) from exc
    if not isinstance(data, dict):
        raise ReviewDataError(f"JSONのルートはオブジェクトである必要があります: {path}")
    return data


def _validated_repo_path(repo_root: Path, path: Path, label: str) -> Path:
    repo_root = repo_root.resolve()
    supplied = Path(path)
    if ".." in supplied.parts:
        raise ReviewDataError(f"{label}に`..`は使用できません: {path}")
    lexical = supplied if supplied.is_absolute() else repo_root / supplied
    lexical = Path(os.path.abspath(lexical))
    try:
        relative = lexical.relative_to(repo_root)
    except ValueError as exc:
        raise ReviewDataError(f"{label}はリポジトリルート配下にしてください: {path}") from exc
    current = repo_root
    for part in relative.parts:
        current = current / part
        if _is_link_like(current):
            raise ReviewDataError(
                f"{label}の祖先または対象にsymlink/junctionは使用できません: "
                f"{relative.as_posix()}"
            )
    try:
        lexical.resolve(strict=False).relative_to(repo_root)
    except ValueError as exc:
        raise ReviewDataError(f"{label}がリポジトリルート外へ解決されます: {path}") from exc
    return lexical


def _relative_manifest_path(
    repo_root: Path, value: Any, expected_prefix: str, label: str
) -> str:
    if not isinstance(value, str) or not value.strip():
        raise ReviewDataError(f"{label}がありません")
    supplied = Path(value)
    if supplied.is_absolute() or ".." in supplied.parts:
        raise ReviewDataError(f"{label}はroot相対パスにしてください: {value}")
    candidate = _validated_repo_path(repo_root, repo_root / supplied, label)
    relative = candidate.relative_to(repo_root.resolve()).as_posix()
    prefix = expected_prefix.rstrip("/") + "/"
    if relative != expected_prefix.rstrip("/") and not relative.startswith(prefix):
        raise ReviewDataError(f"{label}は`{expected_prefix}/`配下にしてください: {value}")
    return relative


def _ensure_new_output(
    repo_root: Path, output: Path, allowed: set[Path], label: str
) -> Path:
    candidate = _validated_repo_path(repo_root, output, label)
    allowed_lexical = {
        _validated_repo_path(repo_root, path, f"{label}許可先") for path in allowed
    }
    if candidate not in allowed_lexical:
        raise ReviewDataError(f"{label}は所定のrun-id別出力先へ指定してください: {output}")
    if candidate.exists() or _is_link_like(candidate):
        raise ReviewDataError(f"既存の{label}は上書きしません: {candidate.name}")
    return candidate


def _atomic_write_text(repo_root: Path, output: Path, text: str) -> None:
    _validated_repo_path(repo_root, output, "出力先")
    output.parent.mkdir(parents=True, exist_ok=True)
    temporary = output.with_name(f".{output.name}.{uuid.uuid4().hex}.tmp")
    _validated_repo_path(repo_root, temporary, "一時出力先")
    try:
        with temporary.open("x", encoding="utf-8") as stream:
            stream.write(text)
            stream.flush()
            os.fsync(stream.fileno())
        try:
            os.link(temporary, output)
        except FileExistsError as exc:
            raise ReviewDataError(f"既存出力は上書きしません: {output.name}") from exc
        except OSError as exc:
            raise ReviewDataError(f"出力を原子的に確定できません: {output.name} ({exc})") from exc
    finally:
        try:
            temporary.unlink()
        except FileNotFoundError:
            pass


def _normalize_entry(entry: Any, role: str) -> dict[str, Any]:
    if isinstance(entry, str):
        entry = {"path": entry, "markdown": entry}
    if not isinstance(entry, dict):
        raise ReviewDataError(f"{role}の要素はオブジェクトである必要があります")
    result = dict(entry)
    if not isinstance(result.get("path"), str) or not result["path"].strip():
        raise ReviewDataError(f"{role}のpathがありません")
    if not isinstance(result.get("markdown"), str) or not result["markdown"].strip():
        raise ReviewDataError(f"{role}のmarkdownがありません: {result['path']}")
    if role == "checklists":
        header_row = result.get("header_row", 1)
        if not isinstance(header_row, int) or header_row < 1:
            raise ReviewDataError(f"header_rowは1以上の整数です: {result['path']}")
        result["header_row"] = header_row
    return result


def _repo_root_from_manifest(manifest_path: Path) -> Path:
    lexical = Path(os.path.abspath(manifest_path))
    if (
        lexical.name != "manifest.json"
        or not re.fullmatch(r"\d{12}", lexical.parent.name)
        or lexical.parent.parent.name != "review-runs"
        or lexical.parent.parent.parent.name != "work"
    ):
        raise ReviewDataError(
            "manifestはwork/review-runs/<yyyyMMddhhmm>/manifest.jsonに配置してください"
        )
    repo_root = lexical.parents[3].resolve()
    _validated_repo_path(repo_root, lexical, "manifest")
    return repo_root


def _repo_root_from_results(results_path: Path) -> tuple[Path, str, Path]:
    lexical = Path(os.path.abspath(results_path))
    if (
        lexical.name != "results.json"
        or not re.fullmatch(r"\d{12}", lexical.parent.name)
        or lexical.parent.parent.name != "review-runs"
        or lexical.parent.parent.parent.name != "work"
    ):
        raise ReviewDataError(
            "resultsはwork/review-runs/<yyyyMMddhhmm>/results.jsonを指定してください"
        )
    repo_root = lexical.parents[3].resolve()
    candidate = _validated_repo_path(repo_root, lexical, "results")
    return repo_root, lexical.parent.name, candidate


def _markdown_text(repo_root: Path, markdown: str) -> str:
    relative = Path(markdown)
    if relative.is_absolute():
        raise ReviewDataError(f"Markdownパスはリポジトリ相対で指定してください: {markdown}")
    candidate = _validated_repo_path(repo_root, repo_root / relative, "Markdown")
    try:
        return candidate.read_text(encoding="utf-8-sig")
    except FileNotFoundError as exc:
        raise ReviewDataError(f"Markdownが見つかりません: {markdown}") from exc


def _split_table_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    return [cell.strip().replace("\\|", "|") for cell in stripped.split("|")]


def _is_separator_row(cells: list[str]) -> bool:
    return bool(cells) and all(
        re.fullmatch(r":?-{3,}:?", cell.replace(" ", "")) for cell in cells
    )


def _check_item_text(cells: list[str]) -> str:
    meaningful = [cell for cell in cells if cell]
    if not meaningful:
        return ""
    for cell in meaningful:
        if not re.fullmatch(r"(?:\d+|[A-Za-z]?\d+(?:\.\d+)*)", cell):
            return cell
    return " / ".join(meaningful)


def _extract_checklist_items(
    checklist: dict[str, Any], markdown_text: str
) -> list[dict[str, Any]]:
    """Extract rows from ordinary Markdown tables, preserving a useful source row."""
    lines = markdown_text.splitlines()
    sheet = ""
    items: list[dict[str, Any]] = []
    table_source_row = checklist["header_row"]
    in_table = False
    for index, line in enumerate(lines):
        heading = re.match(
            r"^#{1,6}\s+(?:Sheet|シート)\s*[:：]?\s*(.+?)\s*$", line, re.I
        )
        if heading:
            sheet = heading.group(1).strip()
            in_table = False
            table_source_row = checklist["header_row"]
            continue
        if "|" not in line:
            in_table = False
            continue
        cells = _split_table_row(line)
        next_cells = (
            _split_table_row(lines[index + 1])
            if index + 1 < len(lines) and "|" in lines[index + 1]
            else []
        )
        if _is_separator_row(cells):
            in_table = True
            table_source_row = checklist["header_row"]
            continue
        if not in_table and _is_separator_row(next_cells):
            continue
        if not in_table:
            continue
        table_source_row += 1
        text = _check_item_text(cells)
        if text:
            items.append(
                {
                    "checklist_file": checklist["path"],
                    "sheet": sheet,
                    "row": table_source_row,
                    "check_item": text,
                }
            )
    return items


def _validate_manifest_paths(
    repo_root: Path,
    run_id: str,
    manifest: dict[str, Any],
    collections: tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]],
) -> None:
    if manifest.get("root") != ".":
        raise ReviewDataError("manifestのrootは`.`である必要があります")
    if manifest.get("run_id") != run_id:
        raise ReviewDataError("manifestのrun_idが配置フォルダと一致しません")
    checklist_paths: set[str] = set()
    for role, input_prefix, entries in zip(
        ("checklists", "references", "targets"),
        ("input/checklists", "input/references", "input/targets"),
        collections,
    ):
        for index, entry in enumerate(entries):
            context = f"{role}[{index}]"
            unknown_path_keys = {
                key
                for key in entry
                if key.endswith("_path")
                and key not in {"original_path", "intermediate_path"}
            }
            if unknown_path_keys:
                raise ReviewDataError(
                    f"{context}に未対応のパス項目があります: "
                    + ", ".join(sorted(unknown_path_keys))
                )
            source = _relative_manifest_path(
                repo_root, entry.get("path"), input_prefix, f"{context}.path"
            )
            original = _relative_manifest_path(
                repo_root,
                entry.get("original_path"),
                input_prefix,
                f"{context}.original_path",
            )
            if source != original:
                raise ReviewDataError(f"{context}.pathとoriginal_pathが一致しません")
            if role == "checklists":
                checklist_paths.add(source)
            _relative_manifest_path(
                repo_root,
                entry.get("markdown"),
                f"work/markdown/{run_id}/{role}",
                f"{context}.markdown",
            )
            intermediate = entry.get("intermediate_path")
            if intermediate is not None:
                _relative_manifest_path(
                    repo_root,
                    intermediate,
                    f"work/converted-office/{run_id}/{role}",
                    f"{context}.intermediate_path",
                )
    checklist_items = manifest.get("checklist_items")
    if not isinstance(checklist_items, list):
        raise ReviewDataError("manifestのchecklist_itemsは配列である必要があります")
    for index, item in enumerate(checklist_items):
        if not isinstance(item, dict):
            raise ReviewDataError(f"checklist_items[{index}]が不正です")
        checklist_file = _relative_manifest_path(
            repo_root,
            item.get("checklist_file"),
            "input/checklists",
            f"checklist_items[{index}].checklist_file",
        )
        if checklist_file not in checklist_paths:
            raise ReviewDataError(
                f"checklist_items[{index}].checklist_fileがchecklistsにありません"
            )


def prepare_bundle(manifest_path: Path) -> dict[str, Any]:
    manifest = _read_json(manifest_path)
    repo_root = _repo_root_from_manifest(manifest_path)
    run_id = manifest.get("run_id")
    if not isinstance(run_id, str) or not run_id:
        raise ReviewDataError("manifestにrun_idがありません")
    generated_at = manifest.get("generated_at") or datetime.now().astimezone().isoformat(
        timespec="seconds"
    )
    checklists = [
        _normalize_entry(item, "checklists") for item in manifest.get("checklists", [])
    ]
    references = [
        _normalize_entry(item, "references") for item in manifest.get("references", [])
    ]
    targets = [_normalize_entry(item, "targets") for item in manifest.get("targets", [])]
    _validate_manifest_paths(
        repo_root, run_id, manifest, (checklists, references, targets)
    )
    if not checklists:
        raise ReviewDataError("チェックリストがありません")
    if not targets:
        raise ReviewDataError("レビュー対象がありません")

    supplied_items = manifest.get("checklist_items")
    checklist_items: list[dict[str, Any]] = []
    materials: list[dict[str, str]] = []
    for role, entries in (
        ("checklist", checklists),
        ("reference", references),
        ("target", targets),
    ):
        for entry in entries:
            content = _markdown_text(repo_root, entry["markdown"])
            materials.append(
                {
                    "role": role,
                    "path": entry["path"],
                    "markdown": entry["markdown"],
                    "content": content,
                }
            )
            if role == "checklist" and supplied_items is None:
                checklist_items.extend(_extract_checklist_items(entry, content))

    if supplied_items is not None:
        if not isinstance(supplied_items, list):
            raise ReviewDataError("checklist_itemsは配列である必要があります")
        checklist_items = [dict(item) for item in supplied_items if isinstance(item, dict)]
        if len(checklist_items) != len(supplied_items):
            raise ReviewDataError("checklist_itemsの要素はオブジェクトである必要があります")
    if not checklist_items:
        raise ReviewDataError(
            "チェック項目を抽出できません。manifestのchecklist_itemsへ行情報を明示してください"
        )

    items: list[dict[str, Any]] = []
    for checklist_item in checklist_items:
        for target in targets:
            items.append(
                {
                    **checklist_item,
                    "target_file": target["path"],
                    "result": "",
                    "comment": "",
                    "evidence": [],
                    "improvement": "",
                }
            )

    return {
        "run_id": run_id,
        "generated_at": generated_at,
        "checklists": checklists,
        "references": references,
        "targets": targets,
        "checklist_items": checklist_items,
        "items": items,
        "review_instructions": {
            "allowed_results": list(RESULT_VALUES),
            "complete_every_combination": True,
            "do_not_infer_applicability": True,
            "evidence_location_required_in_comment": True,
            "improvement_required_for_nonconformity": True,
        },
        "materials": materials,
    }


def _require_string(
    record: dict[str, Any], key: str, context: str, errors: list[str]
) -> str:
    value = record.get(key)
    if not isinstance(value, str) or not value.strip():
        errors.append(f"{context}: {key}は空でない文字列が必要です")
        return ""
    return value.strip()


def _item_key(record: dict[str, Any]) -> tuple[Any, ...]:
    values: list[Any] = []
    for field in ITEM_KEY_FIELDS:
        value = record.get(field)
        try:
            hash(value)
        except TypeError:
            value = json.dumps(value, ensure_ascii=False, sort_keys=True)
        values.append(value)
    return tuple(values)


def validate_results(data: dict[str, Any]) -> list[str]:
    errors: list[str] = []
    _require_string(data, "run_id", "ルート", errors)
    _require_string(data, "generated_at", "ルート", errors)

    normalized: dict[str, list[dict[str, Any]]] = {}
    for collection in ("checklists", "references", "targets"):
        raw = data.get(collection)
        if not isinstance(raw, list):
            errors.append(f"{collection}は配列である必要があります")
            normalized[collection] = []
            continue
        normalized[collection] = []
        for index, entry in enumerate(raw):
            if not isinstance(entry, dict):
                errors.append(f"{collection}[{index}]はオブジェクトである必要があります")
                continue
            _require_string(entry, "path", f"{collection}[{index}]", errors)
            _require_string(entry, "markdown", f"{collection}[{index}]", errors)
            normalized[collection].append(entry)
    if not normalized.get("checklists"):
        errors.append("checklistsは1件以上必要です")
    if not normalized.get("targets"):
        errors.append("targetsは1件以上必要です")

    checklist_items = data.get("checklist_items")
    if not isinstance(checklist_items, list):
        errors.append("checklist_itemsは配列である必要があります")
        checklist_items = []
    expected: dict[tuple[Any, ...], dict[str, Any]] = {}
    target_paths = [entry.get("path") for entry in normalized.get("targets", [])]
    checklist_paths = {entry.get("path") for entry in normalized.get("checklists", [])}
    for index, item in enumerate(checklist_items):
        context = f"checklist_items[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{context}はオブジェクトである必要があります")
            continue
        checklist_file = _require_string(item, "checklist_file", context, errors)
        sheet = item.get("sheet")
        if not isinstance(sheet, str):
            errors.append(f"{context}: sheetは文字列である必要があります")
            sheet = ""
        row = item.get("row")
        if not isinstance(row, int) or row < 1:
            errors.append(f"{context}: rowは1以上の整数である必要があります")
        check_item = _require_string(item, "check_item", context, errors)
        if checklist_file and checklist_file not in checklist_paths:
            errors.append(f"{context}: 未登録のchecklist_fileです: {checklist_file}")
        for target_path in target_paths:
            key = _item_key(
                {
                    "checklist_file": checklist_file,
                    "sheet": sheet,
                    "row": row,
                    "target_file": target_path,
                }
            )
            if key in expected:
                errors.append(f"{context}: チェック項目が重複しています: {key}")
            expected[key] = {"check_item": check_item}

    raw_items = data.get("items")
    if not isinstance(raw_items, list):
        errors.append("itemsは配列である必要があります")
        raw_items = []
    actual: dict[tuple[Any, ...], dict[str, Any]] = {}
    for index, item in enumerate(raw_items):
        context = f"items[{index}]"
        if not isinstance(item, dict):
            errors.append(f"{context}はオブジェクトである必要があります")
            continue
        key = _item_key(item)
        if key in actual:
            errors.append(f"{context}: 組合せが重複しています: {key}")
        actual[key] = item
        if key not in expected:
            errors.append(f"{context}: 未定義の組合せです: {key}")
        expected_check_item = expected.get(key, {}).get("check_item")
        check_item = _require_string(item, "check_item", context, errors)
        if expected_check_item and check_item != expected_check_item:
            errors.append(f"{context}: check_itemがchecklist_itemsと一致しません")
        result = item.get("result")
        if result not in RESULT_VALUES:
            errors.append(f"{context}: resultは{', '.join(RESULT_VALUES)}のいずれかです")
        comment = _require_string(item, "comment", context, errors)
        evidence = item.get("evidence")
        if not isinstance(evidence, list) or not evidence:
            errors.append(f"{context}: evidenceは1件以上の配列が必要です")
            evidence_strings: list[str] = []
        else:
            evidence_strings = []
            for evidence_index, location in enumerate(evidence):
                if not isinstance(location, str) or not location.strip():
                    errors.append(
                        f"{context}: evidence[{evidence_index}]は空でない文字列が必要です"
                    )
                else:
                    evidence_strings.append(location.strip())
        if comment and evidence_strings and not any(
            location in comment for location in evidence_strings
        ):
            errors.append(f"{context}: comment内にevidenceの証拠位置を1件以上記載してください")
        improvement = item.get("improvement", "")
        if not isinstance(improvement, str):
            errors.append(f"{context}: improvementは文字列である必要があります")
        elif result == "不適合" and not improvement.strip():
            errors.append(f"{context}: 不適合には具体的なimprovementが必要です")

    for key in sorted(set(expected) - set(actual), key=str):
        errors.append(f"itemsに不足している組合せがあります: {key}")
    return errors


def _md(value: Any) -> str:
    return str(value).replace("|", "\\|").replace("\n", "<br>")


def _file_list(entries: list[dict[str, Any]]) -> list[str]:
    return [f"- `{entry['path']}`" for entry in entries] or ["- なし"]


def render_summary(data: dict[str, Any]) -> str:
    errors = validate_results(data)
    if errors:
        raise ReviewDataError("結果JSONが不正です:\n- " + "\n- ".join(errors))
    counts = Counter(item["result"] for item in data["items"])
    lines = [
        "# 設計書レビューサマリー",
        "",
        "## 実行情報",
        "",
        f"- 実行ID: `{data['run_id']}`",
        f"- 生成日時: `{data['generated_at']}`",
        "",
        "## 使用ファイル",
        "",
        "### チェックリスト",
        "",
        *_file_list(data["checklists"]),
        "",
        "### 参考資料",
        "",
        *_file_list(data["references"]),
        "",
        "### レビュー対象",
        "",
        *_file_list(data["targets"]),
        "",
        "## 判定件数",
        "",
        "| 判定 | 件数 |",
        "|---|---:|",
    ]
    lines.extend(f"| {result} | {counts[result]} |" for result in RESULT_VALUES)

    nonconformities = [item for item in data["items"] if item["result"] == "不適合"]
    lines.extend(["", "## 不適合項目と改善案", ""])
    if not nonconformities:
        lines.append("不適合なし")
    else:
        for number, item in enumerate(nonconformities, 1):
            location = (
                f"{item['checklist_file']} / "
                f"{item['sheet'] or '(シート指定なし)'}!{item['row']}"
            )
            lines.extend(
                [
                    f"### {number}. {_md(item['check_item'])}",
                    "",
                    f"- 対象ファイル: `{item['target_file']}`",
                    f"- チェックリスト位置: `{location}`",
                    f"- 判定根拠: {_md(item['comment'])}",
                    "- 証拠位置:",
                    *[f"  - `{_md(evidence)}`" for evidence in item["evidence"]],
                    f"- 改善案: {_md(item['improvement'])}",
                    "",
                ]
            )

    confirmations = [item for item in data["items"] if item["result"] == "要確認"]
    lines.extend(["", "## 要確認", ""])
    if not confirmations:
        lines.append("要確認なし")
    else:
        lines.extend(
            [
                "| チェック項目 | 対象ファイル | チェックリスト位置 | 確認が必要な内容 | 証拠位置 |",
                "|---|---|---|---|---|",
            ]
        )
        for item in confirmations:
            location = (
                f"{item['checklist_file']} / "
                f"{item['sheet'] or '(シート指定なし)'}!{item['row']}"
            )
            evidence = "<br>".join(_md(value) for value in item["evidence"])
            lines.append(
                "| "
                + " | ".join(
                    [
                        _md(item["check_item"]),
                        f"`{_md(item['target_file'])}`",
                        f"`{_md(location)}`",
                        _md(item["comment"]),
                        evidence,
                    ]
                )
                + " |"
            )
    return "\n".join(lines).rstrip() + "\n"


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Markdown設計書レビューの準備・検証・集計")
    subparsers = parser.add_subparsers(dest="command", required=True)
    prepare = subparsers.add_parser("prepare", help="manifestからレビューbundleを作る")
    prepare.add_argument("--manifest", type=Path, required=True)
    prepare.add_argument("--output", type=Path, required=True)
    validate = subparsers.add_parser("validate", help="results JSONを検証する")
    validate.add_argument("--results", type=Path, required=True)
    summary = subparsers.add_parser("summary", help="results JSONからsummary.mdを作る")
    summary.add_argument("--results", type=Path, required=True)
    summary.add_argument("--output", type=Path, required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        if args.command == "prepare":
            manifest_path = Path(os.path.abspath(args.manifest))
            repo_root = _repo_root_from_manifest(manifest_path)
            run_id = manifest_path.parent.name
            output = _ensure_new_output(
                repo_root,
                args.output,
                {repo_root / "work" / "review-runs" / run_id / "review_bundle.json"},
                "レビューbundle",
            )
            bundle = prepare_bundle(manifest_path)
            _atomic_write_text(
                repo_root,
                output,
                json.dumps(bundle, ensure_ascii=False, indent=2) + "\n",
            )
            print(f"レビューbundleを作成しました: {args.output}")
            return 0
        repo_root, run_id, results_path = _repo_root_from_results(args.results)
        data = _read_json(results_path)
        if args.command == "validate":
            errors = validate_results(data)
            if errors:
                print("レビュー結果の検証に失敗しました:", file=sys.stderr)
                for error in errors:
                    print(f"- {error}", file=sys.stderr)
                return 1
            print("レビュー結果は有効です。")
            return 0
        output = _ensure_new_output(
            repo_root,
            args.output,
            {
                repo_root
                / "work"
                / "review-runs"
                / run_id
                / "finalize-staging"
                / "summary.md",
                repo_root / "output" / "reviews" / run_id / "summary.md",
            },
            "summary.md",
        )
        summary_text = render_summary(data)
        _atomic_write_text(repo_root, output, summary_text)
        print(f"サマリーを作成しました: {args.output}")
        return 0
    except ReviewDataError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
