#!/usr/bin/env python3
"""画像文書のメタデータ・OCRを Markdown 化し、非破壊で限定編集する。"""

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
from urllib.parse import quote


ROLES = ("checklist", "reference", "target")
SUPPORTED_SUFFIXES = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
FORMAT_BY_SUFFIX = {
    ".png": "PNG",
    ".jpg": "JPEG",
    ".jpeg": "JPEG",
    ".tif": "TIFF",
    ".tiff": "TIFF",
    ".bmp": "BMP",
    ".webp": "WEBP",
}
MULTIFRAME_FORMATS = {"TIFF", "WEBP"}


class ImageDocumentError(RuntimeError):
    """利用者が修正できる画像処理エラー。"""


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
            raise ImageDocumentError(
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
            raise ImageDocumentError(
                f"出力先がシンボリックリンクです。リンク切れも上書きしません: {path}"
            )
        raise ImageDocumentError(f"出力先は既に存在します。上書きしません: {path}")


def _prepare_output_directory(path: Path) -> Path:
    directory = _absolute_without_resolving(path)
    _reject_link_ancestors(directory)
    if _lexists(directory):
        if _is_link_like(directory):
            raise ImageDocumentError(f"画像出力先がシンボリックリンクです: {directory}")
        if not directory.is_dir():
            raise ImageDocumentError(f"画像出力先がフォルダではありません: {directory}")
    else:
        directory.mkdir(parents=True)
    _reject_link_ancestors(directory)
    if _is_link_like(directory):
        raise ImageDocumentError(f"画像出力先がシンボリックリンクです: {directory}")
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
        raise ImageDocumentError(
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


def _save_image_new(image: Any, path: Path, image_format: str, **options: Any) -> None:
    temporary = _temporary_output(path, suffix=path.suffix)
    try:
        image.save(temporary, format=image_format, **options)
        _commit_temporary(temporary, path)
    finally:
        _remove_file(temporary)


def _require_python_version() -> None:
    if sys.version_info < (3, 12):
        raise ImageDocumentError(
            f"Python 3.12 以上が必要です。現在のバージョン: {sys.version.split()[0]}"
        )


def _load_pillow() -> tuple[Any, Any, Any]:
    try:
        from PIL import ExifTags, Image, ImageOps
    except ImportError as exc:  # pragma: no cover - 環境依存
        raise ImageDocumentError(
            "画像処理に必要な Pillow を import できません。README の手順を確認してください。"
        ) from exc
    return Image, ImageOps, ExifTags


def _require_input(path: Path) -> Path:
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise ImageDocumentError(f"入力ファイルが見つかりません: {path}")
    if resolved.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ImageDocumentError(f"対応していない画像形式です: {resolved.suffix}")
    return resolved


def _prepare_output(input_path: Path, output_path: Path) -> Path:
    output = _absolute_without_resolving(output_path)
    if output == input_path:
        raise ImageDocumentError("入力ファイルを上書きできません。別の出力先を指定してください。")
    _assert_new_output_path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    _reject_link_ancestors(output)
    return output


def _relative_source(input_path: Path, repo_root: Path) -> str:
    root = repo_root.expanduser().resolve()
    if not root.is_dir():
        raise ImageDocumentError(f"リポジトリルートが見つかりません: {repo_root}")
    try:
        return input_path.relative_to(root).as_posix()
    except ValueError as exc:
        raise ImageDocumentError(
            f"入力ファイルはリポジトリルート配下に配置してください: {input_path}"
        ) from exc


def _yaml_string(value: object) -> str:
    return json.dumps(str(value), ensure_ascii=False)


def _md_cell(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\\", "\\\\").replace("|", "\\|").replace("\r", "").replace("\n", "<br>")


def _md_alt(value: object) -> str:
    return (
        str(value)
        .replace("\\", "\\\\")
        .replace("[", "\\[")
        .replace("]", "\\]")
        .replace("\r", " ")
        .replace("\n", " ")
    )


def _relative_markdown_link(target: Path, markdown_output: Path) -> str:
    try:
        relative = os.path.relpath(target.resolve(), markdown_output.parent.resolve())
    except ValueError as exc:
        raise ImageDocumentError(
            "Markdown と画像が異なるドライブにあるため相対リンクを作成できません。"
        ) from exc
    posix_relative = relative.replace("\\", "/")
    return quote(posix_relative, safe="/:@-._~")


def _safe_metadata_value(value: object) -> str:
    if isinstance(value, bytes):
        return f"<binary {len(value)} bytes>"
    text = str(value)
    return text if len(text) <= 500 else text[:500] + "…"


def _load_pytesseract(lang: str) -> Any:
    try:
        import pytesseract
    except ImportError as exc:  # pragma: no cover - 環境依存
        raise ImageDocumentError(
            "OCRに必要な pytesseract を import できません。README の手順を確認してください。"
        ) from exc
    try:
        pytesseract.get_tesseract_version()
        available = set(pytesseract.get_languages(config=""))
    except Exception as exc:  # pragma: no cover - 環境依存
        raise ImageDocumentError(
            "PATH 上の Tesseract を実行できません。README の手順を確認してください。"
        ) from exc
    required = {item for item in lang.split("+") if item}
    missing = sorted(required - available)
    if missing:
        raise ImageDocumentError(
            "Tesseract の言語データが不足しています: " + ", ".join(missing)
        )
    return pytesseract


def _ocr_rows(image: Any, pytesseract: Any, lang: str) -> list[dict[str, object]]:
    data = pytesseract.image_to_data(
        image,
        lang=lang,
        output_type=pytesseract.Output.DICT,
        config="--psm 6",
    )
    rows: list[dict[str, object]] = []
    for index, raw_text in enumerate(data.get("text", [])):
        text = str(raw_text).strip()
        try:
            confidence = float(data["conf"][index])
        except (TypeError, ValueError):
            confidence = -1.0
        if not text or confidence < 0:
            continue
        rows.append(
            {
                "text": text,
                "confidence": confidence,
                "x": int(data["left"][index]),
                "y": int(data["top"][index]),
                "width": int(data["width"][index]),
                "height": int(data["height"][index]),
            }
        )
    return rows


def _read_frames(source: Path) -> tuple[list[Any], dict[str, object], dict[str, object], str]:
    Image, _, ExifTags = _load_pillow()
    frames: list[Any] = []
    metadata: dict[str, object] = {}
    exif_metadata: dict[str, object] = {}
    try:
        with Image.open(source) as image:
            source_format = image.format or FORMAT_BY_SUFFIX[source.suffix.lower()]
            metadata = dict(image.info)
            try:
                exif = image.getexif()
                for key, value in exif.items():
                    exif_metadata[str(ExifTags.TAGS.get(key, key))] = value
            except Exception:
                exif_metadata = {}
            frame_count = int(getattr(image, "n_frames", 1))
            for frame_index in range(frame_count):
                image.seek(frame_index)
                frames.append(image.copy())
    except Exception as exc:
        raise ImageDocumentError(f"画像を読み取れません: {source} ({exc})") from exc
    return frames, metadata, exif_metadata, source_format


def image_to_markdown(
    input_path: Path,
    output_path: Path,
    *,
    role: str,
    repo_root: Path,
    images_dir: Path | None = None,
    ocr: bool = False,
    lang: str = "jpn+eng",
) -> Path:
    """画像の全フレームを座標付き OCR 対応 Markdown へ変換する。"""
    if role not in ROLES:
        raise ImageDocumentError(f"role は {', '.join(ROLES)} のいずれかを指定してください。")
    source = _require_input(input_path)
    output = _prepare_output(source, output_path)
    if output.suffix.lower() not in {".md", ".markdown"}:
        raise ImageDocumentError("Markdown の出力拡張子は .md または .markdown にしてください。")
    source_relative = _relative_source(source, repo_root)
    frames, metadata, exif_metadata, source_format = _read_frames(source)
    if not frames:
        raise ImageDocumentError("画像に読み取れるフレームがありません。")
    first_width, first_height = frames[0].size
    source_extension = source.suffix.lstrip(".").lower()
    ocr_languages = [item for item in lang.split("+") if item] if ocr else []
    pytesseract = _load_pytesseract(lang) if ocr else None

    frame_output_dir: Path | None = None
    frame_paths: list[Path] = []
    if images_dir is not None:
        frame_output_dir = _prepare_output_directory(images_dir)
        frame_paths = [
            frame_output_dir / f"{source.stem}-frame-{index:04d}.png"
            for index in range(1, len(frames) + 1)
        ]
        for frame_path in frame_paths:
            _assert_new_output_path(frame_path)

    lines = [
        "---",
        f"source_path: {_yaml_string(source_relative)}",
        f"source_name: {_yaml_string(source.name)}",
        f"source_format: {_yaml_string(source_extension)}",
        f"document_role: {_yaml_string(role)}",
        f"image_width_px: {first_width}",
        f"image_height_px: {first_height}",
        f"ocr_executed: {'true' if ocr else 'false'}",
    ]
    if ocr_languages:
        lines.append("ocr_languages:")
        lines.extend(f"  - {_yaml_string(language)}" for language in ocr_languages)
    else:
        lines.append("ocr_languages: []")
    lines.extend(
        [
            f"converted_at: {_yaml_string(datetime.now().astimezone().isoformat(timespec='seconds'))}",
            'converter_skill: "image-document"',
            f"frame_count: {len(frames)}",
            "---",
            "",
            f"# {_md_cell(source.name)}",
            "",
            "## 画像メタデータ",
            "",
            f"- 形式: `{_md_cell(source_format)}`",
            f"- フレーム数: `{len(frames)}`",
            f"- 先頭フレーム寸法: `{first_width} × {first_height}` px",
            "",
        ]
    )
    combined_metadata = {f"info.{key}": value for key, value in metadata.items()}
    combined_metadata.update({f"exif.{key}": value for key, value in exif_metadata.items()})
    if combined_metadata:
        lines.extend(["| 項目 | 値 |", "| --- | --- |"]) 
        for key, value in sorted(combined_metadata.items()):
            lines.append(f"| {_md_cell(key)} | {_md_cell(_safe_metadata_value(value))} |")
    else:
        lines.append("（追加メタデータなし）")

    for index, frame in enumerate(frames, start=1):
        width, height = frame.size
        lines.extend(
            [
                "",
                f"## フレーム {index}",
                "",
                f"- 根拠位置: `frame.{index}`",
                f"- 寸法: `{width} × {height}` px",
                f"- カラーモード: `{_md_cell(frame.mode)}`",
            ]
        )
        link_target = source
        if frame_paths:
            _save_image_new(frame, frame_paths[index - 1], "PNG")
            link_target = frame_paths[index - 1]
            try:
                frame_relative = frame_paths[index - 1].relative_to(repo_root.expanduser().resolve()).as_posix()
            except ValueError:
                frame_relative = str(frame_paths[index - 1])
            lines.append(f"- 抽出フレーム: `{_md_cell(frame_relative)}`")
        frame_link = _relative_markdown_link(link_target, output)
        lines.extend(
            [
                "",
                f"![{_md_alt(source.name)} フレーム {index}]({frame_link})",
            ]
        )

        lines.extend(["", "### OCR", ""])
        if ocr and pytesseract is not None:
            rows = _ocr_rows(frame, pytesseract, lang)
            if rows:
                lines.extend(
                    [
                        "| 文字列 | 信頼度 | x | y | 幅 | 高さ | 根拠位置 |",
                        "| --- | ---: | ---: | ---: | ---: | ---: | --- |",
                    ]
                )
                for row in rows:
                    evidence = (
                        f"frame.{index} x={row['x']} y={row['y']} "
                        f"w={row['width']} h={row['height']}"
                    )
                    lines.append(
                        f"| {_md_cell(row['text'])} | {float(row['confidence']):.2f} | {row['x']} | "
                        f"{row['y']} | {row['width']} | {row['height']} | {_md_cell(evidence)} |"
                    )
            else:
                lines.append("> 要確認: OCR で文字を認識できませんでした。元画像を目視確認してください。")
        else:
            lines.append("OCR は実行していません。")
            lines.append("> 要確認: 画像中の文字をレビューに使う場合は OCR を実行してください。")
        lines.append(
            "> 要確認: 図形、色、線、配置、写真の意味は OCR から推測せず、必要に応じて元画像を目視確認してください。"
        )

    _write_text_new(output, "\n".join(lines).rstrip() + "\n")
    return output


def _load_operations(path: Path) -> list[dict[str, Any]]:
    source = path.expanduser().resolve()
    if not source.is_file():
        raise ImageDocumentError(f"操作定義ファイルが見つかりません: {path}")
    try:
        payload = json.loads(source.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ImageDocumentError(f"操作定義 JSON を読み取れません: {exc}") from exc
    operations = payload.get("operations") if isinstance(payload, dict) else payload
    if not isinstance(operations, list) or not operations:
        raise ImageDocumentError("操作定義は空でない配列、または operations 配列を持つオブジェクトにしてください。")
    if not all(isinstance(operation, dict) for operation in operations):
        raise ImageDocumentError("各操作は JSON オブジェクトで指定してください。")
    return operations


def _frame_numbers(selector: object, count: int) -> list[int]:
    if selector in (None, "all"):
        return list(range(1, count + 1))
    if not isinstance(selector, list) or not selector:
        raise ImageDocumentError("frames は all または 1 始まりのフレーム番号配列で指定してください。")
    numbers: list[int] = []
    for value in selector:
        if isinstance(value, bool) or not isinstance(value, int) or value < 1 or value > count:
            raise ImageDocumentError(f"フレーム番号が範囲外です: {value}")
        if value not in numbers:
            numbers.append(value)
    return numbers


def _positive_int(value: object, name: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value <= 0:
        raise ImageDocumentError(f"{name} は正の整数で指定してください。")
    return value


def edit_image(input_path: Path, output_path: Path, operations_path: Path) -> Path:
    """JSON の限定操作を全フレームまたは指定フレームへ適用する。"""
    source = _require_input(input_path)
    output = _prepare_output(source, output_path)
    if output.suffix.lower() not in SUPPORTED_SUFFIXES:
        raise ImageDocumentError(f"対応していない出力形式です: {output.suffix}")
    operations = _load_operations(operations_path)
    frames, _, _, _ = _read_frames(source)
    Image, ImageOps, _ = _load_pillow()
    resampling = {
        "nearest": Image.Resampling.NEAREST,
        "bilinear": Image.Resampling.BILINEAR,
        "bicubic": Image.Resampling.BICUBIC,
        "lanczos": Image.Resampling.LANCZOS,
    }
    output_format = FORMAT_BY_SUFFIX[output.suffix.lower()]

    for operation in operations:
        name = operation.get("op")
        selected = _frame_numbers(operation.get("frames", "all"), len(frames))
        if name == "crop":
            box = operation.get("box")
            if (
                not isinstance(box, list)
                or len(box) != 4
                or any(isinstance(value, bool) or not isinstance(value, (int, float)) for value in box)
            ):
                raise ImageDocumentError("crop の box は [left, top, right, bottom] で指定してください。")
            left, top, right, bottom = (int(value) for value in box)
            for number in selected:
                width, height = frames[number - 1].size
                if left < 0 or top < 0 or right > width or bottom > height or left >= right or top >= bottom:
                    raise ImageDocumentError(f"crop の範囲が frame.{number} の画像外です。")
                frames[number - 1] = frames[number - 1].crop((left, top, right, bottom))
        elif name == "rotate":
            degrees = operation.get("degrees")
            if isinstance(degrees, bool) or not isinstance(degrees, (int, float)):
                raise ImageDocumentError("rotate の degrees は数値で指定してください。")
            expand = operation.get("expand", True)
            if not isinstance(expand, bool):
                raise ImageDocumentError("rotate の expand は true または false で指定してください。")
            for number in selected:
                frames[number - 1] = frames[number - 1].rotate(float(degrees), expand=expand)
        elif name == "resize":
            raw_width = operation.get("width")
            raw_height = operation.get("height")
            if raw_width is None and raw_height is None:
                raise ImageDocumentError("resize には width または height を指定してください。")
            width = _positive_int(raw_width, "width") if raw_width is not None else None
            height = _positive_int(raw_height, "height") if raw_height is not None else None
            method = str(operation.get("resample", "lanczos")).lower()
            if method not in resampling:
                raise ImageDocumentError(f"未対応の resample です: {method}")
            for number in selected:
                old_width, old_height = frames[number - 1].size
                new_width = width or max(1, round(old_width * int(height) / old_height))
                new_height = height or max(1, round(old_height * int(width) / old_width))
                frames[number - 1] = frames[number - 1].resize(
                    (new_width, new_height), resample=resampling[method]
                )
        elif name == "grayscale":
            for number in selected:
                frames[number - 1] = ImageOps.grayscale(frames[number - 1])
        elif name == "convert":
            mode = operation.get("mode")
            if mode is not None and (not isinstance(mode, str) or not mode):
                raise ImageDocumentError("convert の mode は Pillow の有効なモード文字列で指定してください。")
            requested_format = operation.get("format")
            if requested_format is not None:
                normalized_format = str(requested_format).upper()
                if normalized_format == "JPG":
                    normalized_format = "JPEG"
                if normalized_format not in set(FORMAT_BY_SUFFIX.values()):
                    raise ImageDocumentError(f"未対応の convert format です: {requested_format}")
                if normalized_format != output_format:
                    raise ImageDocumentError(
                        "convert の format と出力ファイルの拡張子が一致しません。"
                    )
            if mode:
                for number in selected:
                    try:
                        frames[number - 1] = frames[number - 1].convert(mode)
                    except ValueError as exc:
                        raise ImageDocumentError(f"画像モードへ変換できません: {mode}") from exc
        else:
            raise ImageDocumentError(f"未対応の画像操作です: {name}")

    if len(frames) > 1 and output_format not in MULTIFRAME_FORMATS:
        raise ImageDocumentError(
            f"{output_format} へ複数フレームを保存できません。TIFF または WebP を指定してください。"
        )
    if output_format == "JPEG":
        for index, frame in enumerate(frames):
            if frame.mode not in ("RGB", "L", "CMYK"):
                frames[index] = frame.convert("RGB")
    if len(frames) == 1:
        _save_image_new(frames[0], output, output_format)
    else:
        _save_image_new(
            frames[0],
            output,
            output_format,
            save_all=True,
            append_images=frames[1:],
        )
    return output


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="画像文書の Markdown 変換、OCR、限定編集を行います。")
    subparsers = parser.add_subparsers(dest="command", required=True)

    markdown = subparsers.add_parser("to-markdown", help="画像を Markdown に変換します。")
    markdown.add_argument("input", type=Path)
    markdown.add_argument("output", type=Path)
    markdown.add_argument("--role", required=True, choices=ROLES)
    markdown.add_argument("--repo-root", required=True, type=Path)
    markdown.add_argument("--images-dir", type=Path)
    markdown.add_argument("--ocr", action="store_true")
    markdown.add_argument("--lang", default="jpn+eng")

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
            result = image_to_markdown(
                args.input,
                args.output,
                role=args.role,
                repo_root=args.repo_root,
                images_dir=args.images_dir,
                ocr=args.ocr,
                lang=args.lang,
            )
        else:
            result = edit_image(args.input, args.output, args.operations)
    except ImageDocumentError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - Pillow/Tesseract 固有例外の最終境界
        print(f"エラー: 画像処理に失敗しました: {exc}", file=sys.stderr)
        return 1
    print(f"完了: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
