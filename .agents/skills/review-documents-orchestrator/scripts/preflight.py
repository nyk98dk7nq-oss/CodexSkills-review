#!/usr/bin/env python3
"""Check runtime dependencies before a repository review starts."""

from __future__ import annotations

import argparse
import importlib
import json
import os
import platform
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path
from typing import Any


SUPPORTED_EXTENSIONS = {
    ".xlsx",
    ".xls",
    ".docx",
    ".doc",
    ".pptx",
    ".ppt",
    ".pdf",
    ".png",
    ".jpg",
    ".jpeg",
    ".tif",
    ".tiff",
    ".bmp",
    ".webp",
}
IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
LEGACY_APPLICATIONS = {
    ".xls": "Excel.Application",
    ".doc": "Word.Application",
    ".ppt": "PowerPoint.Application",
}


def _is_link_like(path: Path) -> bool:
    return path.is_symlink() or path.is_junction()
IMPORT_REQUIREMENTS = {
    ".xlsx": ("openpyxl",),
    ".xls": ("openpyxl", "pythoncom", "win32com.client"),
    ".docx": ("docx",),
    ".doc": ("docx", "pythoncom", "win32com.client"),
    ".pptx": ("pptx",),
    ".ppt": ("pptx", "pythoncom", "win32com.client"),
    ".pdf": ("pypdf", "pdfplumber", "PIL", "pytesseract"),
    ".png": ("PIL", "pytesseract"),
    ".jpg": ("PIL", "pytesseract"),
    ".jpeg": ("PIL", "pytesseract"),
    ".tif": ("PIL", "pytesseract"),
    ".tiff": ("PIL", "pytesseract"),
    ".bmp": ("PIL", "pytesseract"),
    ".webp": ("PIL", "pytesseract"),
}


def inventory_extensions(root: Path) -> tuple[list[str], set[str]]:
    root = root.resolve()

    def validate(path: Path, label: str) -> None:
        lexical = Path(os.path.abspath(path))
        try:
            relative = lexical.relative_to(root)
        except ValueError as exc:
            raise ValueError(f"{label}がリポジトリルート外です") from exc
        current = root
        for part in relative.parts:
            current = current / part
            if _is_link_like(current):
                raise ValueError(
                    f"{label}にsymlink/junctionは使用できません: {relative.as_posix()}"
                )
        try:
            lexical.resolve(strict=False).relative_to(root)
        except ValueError as exc:
            raise ValueError(f"{label}がリポジトリルート外へ解決されます") from exc

    files: list[str] = []
    extensions: set[str] = set()
    for relative_dir in ("input/checklists", "input/references", "input/targets"):
        directory = root / relative_dir
        validate(directory, relative_dir)
        if not directory.is_dir():
            continue
        for path in sorted(directory.rglob("*")):
            validate(path, "入力パス")
            if not path.is_file() or path.name.lower() == "readme.md" or path.name.startswith("."):
                continue
            suffix = path.suffix.lower()
            if suffix in SUPPORTED_EXTENSIONS:
                files.append(path.relative_to(root).as_posix())
                extensions.add(suffix)
    return files, extensions


def _office_inputs_need_ocr(root: Path, files: list[str]) -> bool:
    prefixes = {".xlsx": "xl/media/", ".docx": "word/media/", ".pptx": "ppt/media/"}
    for relative in files:
        path = root / relative
        suffix = path.suffix.lower()
        if suffix in LEGACY_APPLICATIONS:
            # Binary legacy documents cannot be inspected safely before COM conversion.
            return True
        prefix = prefixes.get(suffix)
        if prefix is None:
            continue
        try:
            with zipfile.ZipFile(path) as archive:
                if any(name.startswith(prefix) and not name.endswith("/") for name in archive.namelist()):
                    return True
        except (OSError, zipfile.BadZipFile):
            # The format converter will give the actionable document error later.
            continue
    return False


def _check_imports(modules: set[str]) -> tuple[list[str], list[str]]:
    available: list[str] = []
    missing: list[str] = []
    for module in sorted(modules):
        try:
            importlib.import_module(module)
            available.append(module)
        except (ImportError, OSError) as exc:
            missing.append(f"Pythonライブラリ `{module}` をimportできません: {exc}")
    return available, missing


def _check_tesseract() -> tuple[dict[str, Any], list[str]]:
    detail: dict[str, Any] = {"required_languages": ["jpn", "eng", "jpn_vert"]}
    missing: list[str] = []
    executable = shutil.which("tesseract")
    detail["executable"] = executable
    if not executable:
        missing.append("PATH上にTesseract本体がありません")
        detail["languages"] = []
        return detail, missing
    try:
        completed = subprocess.run(
            [executable, "--list-langs"],
            check=False,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=20,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        missing.append(f"Tesseractの言語データを確認できません: {exc}")
        detail["languages"] = []
        return detail, missing
    languages = {
        line.strip()
        for line in (completed.stdout + "\n" + completed.stderr).splitlines()
        if line.strip() and not line.lower().startswith("list of available languages")
    }
    detail["languages"] = sorted(languages)
    for language in detail["required_languages"]:
        if language not in languages:
            missing.append(f"Tesseract言語データ `{language}` がありません")
    return detail, missing


def _check_office_com(extensions: set[str]) -> tuple[dict[str, Any], list[str]]:
    detail: dict[str, Any] = {}
    missing: list[str] = []
    applications = sorted({LEGACY_APPLICATIONS[ext] for ext in extensions & set(LEGACY_APPLICATIONS)})
    if not applications:
        return detail, missing
    if platform.system() != "Windows":
        missing.append("旧Office形式の変換にはWindows 11とMicrosoft Officeが必要です")
        return {application: False for application in applications}, missing
    try:
        win32_client = importlib.import_module("win32com.client")
    except (ImportError, OSError) as exc:
        missing.append(f"pywin32をimportできません: {exc}")
        return {application: False for application in applications}, missing
    for application in applications:
        instance = None
        try:
            instance = win32_client.DispatchEx(application)
            detail[application] = True
        except Exception as exc:  # COM errors have provider-specific classes.
            detail[application] = False
            missing.append(f"Microsoft Office COM `{application}` を起動できません: {exc}")
        finally:
            if instance is not None:
                try:
                    instance.Quit()
                except Exception:
                    pass
    return detail, missing


def run_checks(root: Path) -> dict[str, Any]:
    root = root.resolve()
    inventory_error: str | None = None
    try:
        files, extensions = inventory_extensions(root)
    except ValueError as exc:
        files, extensions = [], set()
        inventory_error = str(exc)
    missing: list[str] = []
    if inventory_error:
        missing.append(inventory_error)
    if not root.is_dir():
        missing.append("リポジトリルートが見つかりません。--rootの指定を確認してください")
    version_ok = sys.version_info >= (3, 12)
    if not version_ok:
        missing.append(
            f"Python 3.12以上が必要です（実行中: {sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}）"
        )
    modules = {
        module for extension in extensions for module in IMPORT_REQUIREMENTS.get(extension, ())
    }
    office_embedded_images = _office_inputs_need_ocr(root, files)
    if office_embedded_images:
        modules.update({"PIL", "pytesseract"})
    available_imports, import_missing = _check_imports(modules)
    missing.extend(import_missing)

    needs_ocr = bool(extensions & (IMAGE_EXTENSIONS | {".pdf"})) or office_embedded_images
    tesseract: dict[str, Any] | None = None
    if needs_ocr:
        tesseract, tesseract_missing = _check_tesseract()
        missing.extend(tesseract_missing)
    office_com, office_missing = _check_office_com(extensions)
    missing.extend(office_missing)
    return {
        "ok": not missing,
        "root": ".",
        "python": {
            "version": platform.python_version(),
            "minimum": "3.12",
            "ok": version_ok,
        },
        "input_files": files,
        "extensions": sorted(extensions),
        "office_embedded_images": office_embedded_images,
        "imports": {"required": sorted(modules), "available": available_imports},
        "tesseract": tesseract,
        "office_com": office_com,
        "missing": missing,
        "guidance": (
            "不足項目をREADMEの「セットアップ」手順で導入してから再実行してください。"
            if missing
            else "レビュー処理を開始できます。"
        ),
    }


def _print_human(report: dict[str, Any]) -> None:
    print(f"Python: {report['python']['version']} (必要: 3.12以上)")
    print(f"入力ファイル: {len(report['input_files'])}件")
    if report["imports"]["required"]:
        print("確認ライブラリ: " + ", ".join(report["imports"]["required"]))
    if report["ok"]:
        print("環境確認に成功しました。レビュー処理を開始できます。")
    else:
        print("環境確認に失敗しました。処理は開始しません。", file=sys.stderr)
        for item in report["missing"]:
            print(f"- {item}", file=sys.stderr)
        print(report["guidance"], file=sys.stderr)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="設計書レビューの実行前環境確認")
    parser.add_argument("--root", type=Path, required=True)
    parser.add_argument("--json", action="store_true", dest="as_json")
    args = parser.parse_args(argv)
    report = run_checks(args.root)
    if args.as_json:
        print(json.dumps(report, ensure_ascii=False, indent=2))
    else:
        _print_human(report)
    return 0 if report["ok"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
