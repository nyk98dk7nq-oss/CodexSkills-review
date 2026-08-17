#!/usr/bin/env python3
"""PDF を根拠位置付き Markdown へ変換し、限定的に編集する。"""

from __future__ import annotations

import argparse
import errno
import json
import os
import shutil
import sys
import tempfile
from datetime import datetime
from pathlib import Path
from typing import Any, Sequence


ROLES = ("checklist", "reference", "target")
SPARSE_TEXT_THRESHOLD = 20


class PdfDocumentError(RuntimeError):
    """利用者が修正できる PDF 処理エラー。"""


def _absolute_without_resolving(path: Path) -> Path:
    return Path(os.path.abspath(os.fspath(path.expanduser())))


def _lexists(path: Path) -> bool:
    return os.path.lexists(os.fspath(path))


def _is_link_like(path: Path) -> bool:
    if path.is_symlink():
        return True
    is_junction = getattr(path, "is_junction", None)
    if is_junction is None:
        return False
    try:
        return bool(is_junction())
    except OSError:
        return False


def _reject_link_ancestors(path: Path) -> None:
    current = path.parent
    while True:
        if _lexists(current) and _is_link_like(current):
            raise PdfDocumentError(
                f"出力先の祖先フォルダがシンボリックリンクです: {current}"
            )
        parent = current.parent
        if parent == current:
            break
        current = parent


def _assert_new_output_path(path: Path) -> None:
    _reject_link_ancestors(path)
    if _lexists(path):
        if _is_link_like(path):
            raise PdfDocumentError(
                f"出力先がシンボリックリンクです。リンク切れも上書きしません: {path}"
            )
        raise PdfDocumentError(f"出力先は既に存在します。上書きしません: {path}")


def _prepare_output_directory(path: Path) -> Path:
    directory = _absolute_without_resolving(path)
    _reject_link_ancestors(directory)
    if _lexists(directory):
        if _is_link_like(directory):
            raise PdfDocumentError(f"画像出力先がシンボリックリンクです: {directory}")
        if not directory.is_dir():
            raise PdfDocumentError(f"画像出力先がフォルダではありません: {directory}")
    else:
        directory.mkdir(parents=True)
    _reject_link_ancestors(directory)
    if _is_link_like(directory):
        raise PdfDocumentError(f"画像出力先がシンボリックリンクです: {directory}")
    return directory


def _temporary_output(path: Path, suffix: str | None = None) -> Path:
    descriptor, name = tempfile.mkstemp(
        dir=path.parent,
        prefix=f".{path.stem}.",
        suffix=suffix if suffix is not None else path.suffix,
    )
    os.close(descriptor)
    return Path(name)


def _remove_file(path: Path) -> None:
    if _lexists(path):
        os.unlink(path)


def _commit_temporary(temp_path: Path, output_path: Path) -> None:
    _assert_new_output_path(output_path)
    try:
        os.link(temp_path, output_path)
    except FileExistsError as exc:
        raise PdfDocumentError(
            f"出力先は既に存在します。上書きしません: {output_path}"
        ) from exc
    except OSError as exc:
        unsupported = {
            errno.EACCES,
            errno.EPERM,
            errno.EINVAL,
            getattr(errno, "ENOTSUP", -1),
            getattr(errno, "EOPNOTSUPP", -1),
        }
        if exc.errno not in unsupported:
            raise
        created = False
        try:
            with temp_path.open("rb") as source, output_path.open("xb") as destination:
                created = True
                shutil.copyfileobj(source, destination)
                destination.flush()
                os.fsync(destination.fileno())
        except Exception:
            if created:
                _remove_file(output_path)
            raise
    finally:
        _remove_file(temp_path)


def _write_text_new(path: Path, text: str) -> None:
    temporary = _temporary_output(path, suffix=".tmp")
    try:
        temporary.write_text(text, encoding="utf-8", newline="\n")
        _commit_temporary(temporary, path)
    finally:
        _remove_file(temporary)


def _save_image_new(image: Any, path: Path, image_format: str) -> None:
    temporary = _temporary_output(path, suffix=path.suffix)
    try:
        image.save(temporary, format=image_format)
        _commit_temporary(temporary, path)
    finally:
        _remove_file(temporary)


def _require_python_version() -> None:
    if sys.version_info < (3, 12):
        raise PdfDocumentError(
            f"Python 3.12 以上が必要です。現在のバージョン: {sys.version.split()[0]}"
        )


def _load_pdf_libraries() -> tuple[Any, Any, Any]:
    try:
        import pdfplumber
        from pypdf import PdfReader, PdfWriter
    except ImportError as exc:  # pragma: no cover - 環境依存
        raise PdfDocumentError(
            "PDF処理に必要な pypdf と pdfplumber を import できません。"
            "README の手順でライブラリをインストールしてください。"
        ) from exc
    return pdfplumber, PdfReader, PdfWriter


def _require_input(path: Path, suffix: str = ".pdf") -> Path:
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise PdfDocumentError(f"入力ファイルが見つかりません: {path}")
    if resolved.suffix.lower() != suffix:
        raise PdfDocumentError(f"対応していない入力形式です: {resolved.suffix}")
    return resolved


def _prepare_output(input_path: Path, output_path: Path) -> Path:
    output = _absolute_without_resolving(output_path)
    if output == input_path:
        raise PdfDocumentError("入力ファイルを上書きできません。別の出力先を指定してください。")
    _assert_new_output_path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    _reject_link_ancestors(output)
    return output


def _relative_source(input_path: Path, repo_root: Path) -> str:
    root = repo_root.expanduser().resolve()
    if not root.is_dir():
        raise PdfDocumentError(f"リポジトリルートが見つかりません: {repo_root}")
    try:
        return input_path.relative_to(root).as_posix()
    except ValueError as exc:
        raise PdfDocumentError(
            f"入力ファイルはリポジトリルート配下に配置してください: {input_path}"
        ) from exc


def _yaml_string(value: object) -> str:
    return json.dumps(str(value), ensure_ascii=False)


def _md_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\\", "\\\\").replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


def _number(value: object) -> str:
    try:
        return f"{float(value):.2f}"
    except (TypeError, ValueError):
        return ""


def _render_page(pdf_path: Path, page_index: int, plumber_page: Any, dpi: int) -> Any:
    try:
        return plumber_page.to_image(resolution=dpi, antialias=True).original.convert("RGB")
    except Exception as first_error:
        try:
            import pypdfium2 as pdfium

            document = pdfium.PdfDocument(str(pdf_path))
            try:
                bitmap = document[page_index].render(scale=dpi / 72.0)
                return bitmap.to_pil().convert("RGB")
            finally:
                document.close()
        except Exception as second_error:  # pragma: no cover - renderer の構成依存
            raise PdfDocumentError(
                "PDFページを画像化できません。pdfplumber が利用する画像化環境、"
                "または pypdfium2 を確認してください。"
                f" ({first_error}; {second_error})"
            ) from second_error


def _load_pytesseract(lang: str) -> Any:
    try:
        import pytesseract
    except ImportError as exc:  # pragma: no cover - 環境依存
        raise PdfDocumentError(
            "OCRに必要な pytesseract を import できません。README の手順を確認してください。"
        ) from exc
    try:
        pytesseract.get_tesseract_version()
        available = set(pytesseract.get_languages(config=""))
    except Exception as exc:  # pragma: no cover - 環境依存
        raise PdfDocumentError(
            "PATH 上の Tesseract を実行できません。README の手順を確認してください。"
        ) from exc
    required = {item for item in lang.split("+") if item}
    missing = sorted(required - available)
    if missing:
        raise PdfDocumentError(
            "Tesseract の言語データが不足しています: " + ", ".join(missing)
        )
    return pytesseract


def _ocr_rows(image: Any, pytesseract: Any, lang: str, page_width: float, page_height: float) -> list[dict[str, object]]:
    data = pytesseract.image_to_data(
        image,
        lang=lang,
        output_type=pytesseract.Output.DICT,
        config="--psm 6",
    )
    rows: list[dict[str, object]] = []
    image_width, image_height = image.size
    count = len(data.get("text", []))
    for index in range(count):
        text = str(data["text"][index]).strip()
        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError):
            confidence = -1.0
        if not text or confidence < 0:
            continue
        left = int(data["left"][index])
        top = int(data["top"][index])
        width = int(data["width"][index])
        height = int(data["height"][index])
        rows.append(
            {
                "text": text,
                "confidence": confidence,
                "x": left,
                "y": top,
                "width": width,
                "height": height,
                "pdf_x0": left * page_width / image_width,
                "pdf_top": top * page_height / image_height,
                "pdf_x1": (left + width) * page_width / image_width,
                "pdf_bottom": (top + height) * page_height / image_height,
            }
        )
    return rows


def _append_table(lines: list[str], table: Sequence[Sequence[object]]) -> None:
    if not table:
        lines.append("（表データなし）")
        return
    width = max((len(row) for row in table), default=0)
    if width == 0:
        lines.append("（表データなし）")
        return
    normalized = [list(row) + [""] * (width - len(row)) for row in table]
    lines.append("| " + " | ".join(_md_cell(value) for value in normalized[0]) + " |")
    lines.append("| " + " | ".join("---" for _ in range(width)) + " |")
    for row in normalized[1:]:
        lines.append("| " + " | ".join(_md_cell(value) for value in row) + " |")


def pdf_to_markdown(
    input_path: Path,
    output_path: Path,
    *,
    role: str,
    repo_root: Path,
    images_dir: Path | None = None,
    ocr: bool = False,
    ocr_all: bool = False,
    lang: str = "jpn+eng",
    dpi: int = 180,
) -> Path:
    """PDF のテキスト、表、座標、必要な OCR を Markdown に書き出す。"""
    if role not in ROLES:
        raise PdfDocumentError(f"role は {', '.join(ROLES)} のいずれかを指定してください。")
    if dpi < 72 or dpi > 600:
        raise PdfDocumentError("dpi は 72 以上 600 以下で指定してください。")
    source = _require_input(input_path)
    output = _prepare_output(source, output_path)
    if output.suffix.lower() not in {".md", ".markdown"}:
        raise PdfDocumentError("Markdown の出力拡張子は .md または .markdown にしてください。")
    source_relative = _relative_source(source, repo_root)
    image_output: Path | None = None
    if images_dir is not None:
        image_output = _prepare_output_directory(images_dir)
        if image_output == source:
            raise PdfDocumentError("画像出力先に入力ファイルを指定できません。")

    pdfplumber, PdfReader, _ = _load_pdf_libraries()
    reader = PdfReader(str(source))
    if reader.is_encrypted:
        raise PdfDocumentError("暗号化された PDF は読み取れません。暗号化を解除してください。")
    planned_page_images: list[Path] = []
    if image_output is not None:
        planned_page_images = [
            image_output / f"{source.stem}-page-{page_number:04d}.png"
            for page_number in range(1, len(reader.pages) + 1)
        ]
        for planned_image in planned_page_images:
            _assert_new_output_path(planned_image)
    metadata = {str(key): str(value) for key, value in (reader.metadata or {}).items()}
    ocr = ocr or ocr_all
    pytesseract = _load_pytesseract(lang) if ocr else None

    lines = [
        "---",
        f"source_path: {_yaml_string(source_relative)}",
        f"source_name: {_yaml_string(source.name)}",
        'source_format: "pdf"',
        f"document_role: {_yaml_string(role)}",
        f"converted_at: {_yaml_string(datetime.now().astimezone().isoformat(timespec='seconds'))}",
        'converter_skill: "pdf-document"',
        f"page_count: {len(reader.pages)}",
        "---",
        "",
        f"# {_md_cell(source.name)}",
        "",
        "## 文書メタデータ",
        "",
    ]
    if metadata:
        lines.extend(["| 項目 | 値 |", "| --- | --- |"]) 
        for key, value in sorted(metadata.items()):
            lines.append(f"| {_md_cell(key)} | {_md_cell(value)} |")
    else:
        lines.append("（メタデータなし）")

    with pdfplumber.open(str(source)) as document:
        for page_index, page in enumerate(document.pages):
            page_number = page_index + 1
            lines.extend(
                [
                    "",
                    f"## ページ {page_number}",
                    "",
                    f"- 根拠位置: `p.{page_number}`",
                    f"- ページ寸法: `{_number(page.width)} × {_number(page.height)}` pt",
                ]
            )
            text = page.extract_text(x_tolerance=1, y_tolerance=3) or ""
            words = page.extract_words(use_text_flow=True, keep_blank_chars=False) or []
            lines.extend(["", "### 抽出テキスト", ""])
            if text.strip():
                lines.extend(["```text", text.replace("```", "` ` `"), "```"])
            else:
                lines.append("（テキストレイヤーなし）")

            lines.extend(["", "### テキスト座標", ""])
            if words:
                lines.extend(
                    [
                        "| 文字列 | x0 | top | x1 | bottom | 根拠位置 |",
                        "| --- | ---: | ---: | ---: | ---: | --- |",
                    ]
                )
                for word in words:
                    evidence = (
                        f"p.{page_number} x={_number(word.get('x0'))} "
                        f"y={_number(word.get('top'))}"
                    )
                    lines.append(
                        "| "
                        + " | ".join(
                            [
                                _md_cell(word.get("text")),
                                _number(word.get("x0")),
                                _number(word.get("top")),
                                _number(word.get("x1")),
                                _number(word.get("bottom")),
                                _md_cell(evidence),
                            ]
                        )
                        + " |"
                    )
            else:
                lines.append("（座標付きテキストなし）")

            lines.extend(["", "### 表", ""])
            try:
                found_tables = page.find_tables()
            except Exception:
                found_tables = []
            if found_tables:
                for table_index, found in enumerate(found_tables, start=1):
                    bbox = tuple(_number(value) for value in found.bbox)
                    lines.extend(
                        [
                            f"#### 表 {table_index}",
                            "",
                            f"- 根拠位置: `p.{page_number} bbox=({', '.join(bbox)})`",
                            "",
                        ]
                    )
                    _append_table(lines, found.extract())
                    lines.append("")
            else:
                lines.append("（検出された表なし）")

            lines.extend(["", "### PDF 内画像領域", ""])
            if page.images:
                lines.extend(
                    [
                        "| 番号 | x0 | top | x1 | bottom |",
                        "| ---: | ---: | ---: | ---: | ---: |",
                    ]
                )
                for image_index, item in enumerate(page.images, start=1):
                    lines.append(
                        f"| {image_index} | {_number(item.get('x0'))} | {_number(item.get('top'))} | "
                        f"{_number(item.get('x1'))} | {_number(item.get('bottom'))} |"
                    )
            else:
                lines.append("（画像領域なし）")

            compact_text = "".join(text.split())
            should_ocr = ocr and (
                ocr_all
                or len(compact_text) < SPARSE_TEXT_THRESHOLD
                or bool(page.images)
            )
            rendered = None
            if image_output is not None or should_ocr:
                rendered = _render_page(source, page_index, page, dpi)
            if image_output is not None and rendered is not None:
                image_path = planned_page_images[page_index]
                _save_image_new(rendered, image_path, "PNG")
                try:
                    image_relative = image_path.relative_to(repo_root.expanduser().resolve()).as_posix()
                except ValueError:
                    image_relative = str(image_path)
                lines.extend(["", f"- ページ画像: `{_md_cell(image_relative)}`"])

            lines.extend(["", "### OCR", ""])
            if should_ocr and rendered is not None and pytesseract is not None:
                rows = _ocr_rows(rendered, pytesseract, lang, float(page.width), float(page.height))
                if rows:
                    lines.extend(
                        [
                            "| 文字列 | 信頼度 | x | y | 幅 | 高さ | PDF x0 | PDF top | PDF x1 | PDF bottom | 根拠位置 |",
                            "| --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | ---: | --- |",
                        ]
                    )
                    for row in rows:
                        evidence = (
                            f"p.{page_number} OCR x={row['x']} y={row['y']} "
                            f"w={row['width']} h={row['height']}"
                        )
                        lines.append(
                            "| "
                            + " | ".join(
                                [
                                    _md_cell(row["text"]),
                                    _number(row["confidence"]),
                                    str(row["x"]),
                                    str(row["y"]),
                                    str(row["width"]),
                                    str(row["height"]),
                                    _number(row["pdf_x0"]),
                                    _number(row["pdf_top"]),
                                    _number(row["pdf_x1"]),
                                    _number(row["pdf_bottom"]),
                                    _md_cell(evidence),
                                ]
                            )
                            + " |"
                        )
                else:
                    lines.append("> 要確認: OCR で文字を認識できませんでした。元ページを目視確認してください。")
            elif ocr and not should_ocr:
                lines.append("テキストレイヤーが十分にあるため OCR を省略しました。")
            else:
                lines.append("OCR は実行していません。")

            if len(compact_text) < SPARSE_TEXT_THRESHOLD and not should_ocr:
                lines.append("> 要確認: テキストが少ないページです。画像 PDF の可能性があるため OCR または目視確認が必要です。")
            lines.append(
                "> 要確認: 図形、色、線、レイアウトの意味は機械抽出だけで断定せず、必要に応じてページ画像を目視確認してください。"
            )

    _write_text_new(output, "\n".join(lines).rstrip() + "\n")
    return output


def _load_operations(path: Path) -> list[dict[str, Any]]:
    source = path.expanduser().resolve()
    if not source.is_file():
        raise PdfDocumentError(f"操作定義ファイルが見つかりません: {path}")
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise PdfDocumentError(f"操作定義 JSON を読み取れません: {exc}") from exc
    operations = payload.get("operations") if isinstance(payload, dict) else payload
    if not isinstance(operations, list) or not operations:
        raise PdfDocumentError("操作定義は空でない配列、または operations 配列を持つオブジェクトにしてください。")
    if not all(isinstance(operation, dict) for operation in operations):
        raise PdfDocumentError("各操作は JSON オブジェクトで指定してください。")
    return operations


def _page_numbers(selector: object, count: int, *, allow_all: bool = True) -> list[int]:
    if allow_all and selector == "all":
        return list(range(1, count + 1))
    if not isinstance(selector, list) or not selector:
        raise PdfDocumentError("pages は 1 始まりのページ番号配列で指定してください。")
    numbers: list[int] = []
    for value in selector:
        if isinstance(value, bool) or not isinstance(value, int) or value < 1 or value > count:
            raise PdfDocumentError(f"ページ番号が範囲外です: {value}")
        if value not in numbers:
            numbers.append(value)
    return numbers


def edit_pdf(input_path: Path, output_path: Path, operations_path: Path) -> Path:
    """JSON の限定操作だけを順に適用して新しい PDF を保存する。"""
    source = _require_input(input_path)
    output = _prepare_output(source, output_path)
    operations_file = operations_path.expanduser().resolve()
    operations = _load_operations(operations_file)
    _, PdfReader, PdfWriter = _load_pdf_libraries()
    reader = PdfReader(str(source))
    if reader.is_encrypted:
        raise PdfDocumentError("暗号化された PDF は編集できません。暗号化を解除してください。")
    pages = list(reader.pages)
    metadata = {str(key): str(value) for key, value in (reader.metadata or {}).items()}

    for operation in operations:
        name = operation.get("op")
        if name == "rotate":
            degrees = operation.get("degrees")
            if isinstance(degrees, bool) or not isinstance(degrees, int) or degrees % 90 != 0:
                raise PdfDocumentError("rotate の degrees は 90 の倍数で指定してください。")
            for page_number in _page_numbers(operation.get("pages", "all"), len(pages)):
                pages[page_number - 1].rotate(degrees)
        elif name == "delete":
            deleting = set(_page_numbers(operation.get("pages"), len(pages), allow_all=False))
            pages = [page for index, page in enumerate(pages, start=1) if index not in deleting]
            if not pages:
                raise PdfDocumentError("全ページを削除する操作は実行できません。")
        elif name == "reorder":
            order = _page_numbers(operation.get("pages"), len(pages), allow_all=False)
            if sorted(order) != list(range(1, len(pages) + 1)):
                raise PdfDocumentError("reorder の pages は全ページを重複なく1回ずつ指定してください。")
            pages = [pages[number - 1] for number in order]
        elif name == "merge":
            raw_file = operation.get("file")
            if not isinstance(raw_file, str) or not raw_file.strip():
                raise PdfDocumentError("merge には file を指定してください。")
            merge_path = Path(raw_file)
            if not merge_path.is_absolute():
                merge_path = operations_file.parent / merge_path
            merge_source = _require_input(merge_path)
            merged_reader = PdfReader(str(merge_source))
            if merged_reader.is_encrypted:
                raise PdfDocumentError(f"暗号化された PDF は結合できません: {merge_source}")
            selected = _page_numbers(operation.get("pages", "all"), len(merged_reader.pages))
            inserting = [merged_reader.pages[number - 1] for number in selected]
            position = operation.get("position", "end")
            if position == "end":
                insert_index = len(pages)
            elif isinstance(position, int) and not isinstance(position, bool) and 1 <= position <= len(pages) + 1:
                insert_index = position - 1
            else:
                raise PdfDocumentError("merge の position は end または挿入位置（1始まり）で指定してください。")
            pages[insert_index:insert_index] = inserting
        elif name == "metadata":
            values = operation.get("values")
            if not isinstance(values, dict):
                raise PdfDocumentError("metadata の values はオブジェクトで指定してください。")
            for key, value in values.items():
                normalized = str(key)
                if not normalized.startswith("/"):
                    normalized = "/" + normalized
                metadata[normalized] = str(value)
        else:
            raise PdfDocumentError(f"未対応の PDF 操作です: {name}")

    writer = PdfWriter()
    for page in pages:
        writer.add_page(page)
    if metadata:
        writer.add_metadata(metadata)
    temporary = _temporary_output(output, suffix=output.suffix)
    try:
        with temporary.open("wb") as stream:
            writer.write(stream)
            stream.flush()
            os.fsync(stream.fileno())
        _commit_temporary(temporary, output)
    finally:
        _remove_file(temporary)
    return output


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="PDF の Markdown 変換と限定編集を行います。")
    subparsers = parser.add_subparsers(dest="command", required=True)

    markdown = subparsers.add_parser("to-markdown", help="PDF を Markdown に変換します。")
    markdown.add_argument("input", type=Path)
    markdown.add_argument("output", type=Path)
    markdown.add_argument("--role", required=True, choices=ROLES)
    markdown.add_argument("--repo-root", required=True, type=Path)
    markdown.add_argument("--images-dir", type=Path)
    markdown.add_argument("--ocr", action="store_true", help="テキストが少ないページへ OCR を実行します。")
    markdown.add_argument("--ocr-all", action="store_true", help="全ページへ OCR を実行します。")
    markdown.add_argument("--lang", default="jpn+eng")
    markdown.add_argument("--dpi", type=int, default=180)

    edit = subparsers.add_parser("edit", help="JSON で指定した限定操作を適用します。")
    edit.add_argument("input", type=Path)
    edit.add_argument("output", type=Path)
    edit.add_argument("--operations", required=True, type=Path)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        _require_python_version()
        if args.command == "to-markdown":
            result = pdf_to_markdown(
                args.input,
                args.output,
                role=args.role,
                repo_root=args.repo_root,
                images_dir=args.images_dir,
                ocr=args.ocr,
                ocr_all=args.ocr_all,
                lang=args.lang,
                dpi=args.dpi,
            )
        else:
            result = edit_pdf(args.input, args.output, args.operations)
    except PdfDocumentError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - ライブラリ固有例外の最終境界
        print(f"エラー: PDF処理に失敗しました: {exc}", file=sys.stderr)
        return 1
    print(f"完了: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
