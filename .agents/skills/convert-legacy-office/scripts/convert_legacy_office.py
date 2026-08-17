#!/usr/bin/env python3
"""Microsoft Office COM を使い旧 Office 文書を新形式へ非破壊変換する。"""

from __future__ import annotations

import argparse
import errno
import os
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Sequence


class LegacyOfficeError(RuntimeError):
    """利用者が修正できる旧 Office 変換エラー。"""


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
            raise LegacyOfficeError(
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
            raise LegacyOfficeError(
                f"出力先がシンボリックリンクです。リンク切れも上書きしません: {path}"
            )
        raise LegacyOfficeError(f"出力先は既に存在します。上書きしません: {path}")


def _temporary_output(path: Path) -> Path:
    descriptor, name = tempfile.mkstemp(
        dir=path.parent,
        prefix=f".{path.stem}.",
        suffix=path.suffix,
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
        raise LegacyOfficeError(
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


def _require_python_version() -> None:
    if sys.version_info < (3, 12):
        raise LegacyOfficeError(
            f"Python 3.12 以上が必要です。現在のバージョン: {sys.version.split()[0]}"
        )


@dataclass(frozen=True)
class ConversionSpec:
    source_suffix: str
    output_suffix: str
    program_id: str
    file_format: int
    application: str


CONVERSIONS = {
    ".xls": ConversionSpec(".xls", ".xlsx", "Excel.Application", 51, "excel"),
    ".doc": ConversionSpec(".doc", ".docx", "Word.Application", 16, "word"),
    ".ppt": ConversionSpec(".ppt", ".pptx", "PowerPoint.Application", 24, "powerpoint"),
}


def conversion_spec(path: Path | str) -> ConversionSpec:
    suffix = Path(path).suffix.lower()
    try:
        return CONVERSIONS[suffix]
    except KeyError as exc:
        raise LegacyOfficeError(
            f"対応していない旧 Office 形式です: {suffix or '拡張子なし'}"
        ) from exc


def suggested_output(input_path: Path, output_dir: Path) -> Path:
    spec = conversion_spec(input_path)
    return output_dir / input_path.with_suffix(spec.output_suffix).name


def _validate_paths(input_path: Path, output_path: Path) -> tuple[Path, Path, ConversionSpec]:
    source = input_path.expanduser().resolve()
    if not source.is_file():
        raise LegacyOfficeError(f"入力ファイルが見つかりません: {input_path}")
    spec = conversion_spec(source)
    output = _absolute_without_resolving(output_path)
    if output == source:
        raise LegacyOfficeError("入力ファイルを上書きできません。別の出力先を指定してください。")
    if output.suffix.lower() != spec.output_suffix:
        raise LegacyOfficeError(
            f"{source.suffix} の出力拡張子は {spec.output_suffix} にしてください: {output}"
        )
    _assert_new_output_path(output)
    output.parent.mkdir(parents=True, exist_ok=True)
    _reject_link_ancestors(output)
    return source, output, spec


def _load_com() -> tuple[Any, Callable[[str], Any]]:
    try:
        import pythoncom
        from win32com.client import DispatchEx
    except ImportError as exc:  # pragma: no cover - Windows 環境依存
        raise LegacyOfficeError(
            "pywin32 を import できません。README の手順で pywin32 をインストールしてください。"
        ) from exc
    return pythoncom, DispatchEx


def _open_document(application: Any, source: Path, spec: ConversionSpec) -> Any:
    if spec.application == "excel":
        application.Visible = False
        application.DisplayAlerts = False
        return application.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
    if spec.application == "word":
        application.Visible = False
        application.DisplayAlerts = 0
        return application.Documents.Open(
            str(source),
            ConfirmConversions=False,
            ReadOnly=True,
            AddToRecentFiles=False,
        )
    return application.Presentations.Open(
        str(source),
        ReadOnly=True,
        Untitled=False,
        WithWindow=False,
    )


def _save_document(document: Any, output: Path, spec: ConversionSpec) -> None:
    if spec.application == "word":
        document.SaveAs2(str(output), FileFormat=spec.file_format)
    else:
        document.SaveAs(str(output), FileFormat=spec.file_format)


def _close_document(document: Any, spec: ConversionSpec) -> None:
    if spec.application in {"excel", "word"}:
        document.Close(False)
    else:
        document.Close()


def convert_file(
    input_path: Path,
    output_path: Path,
    *,
    platform: str | None = None,
    dispatch_factory: Callable[[str], Any] | None = None,
    pythoncom_module: Any | None = None,
) -> Path:
    """単一の .xls/.doc/.ppt を対応する Open XML 形式へ変換する。"""
    source, output, spec = _validate_paths(input_path, output_path)
    current_platform = sys.platform if platform is None else platform
    if current_platform != "win32":
        raise LegacyOfficeError(
            "旧 Office 形式の変換は Windows と Microsoft Office が必要です。"
            "LibreOffice や外部変換サービスは使用しません。"
        )
    if dispatch_factory is None or pythoncom_module is None:
        loaded_pythoncom, loaded_dispatch = _load_com()
        pythoncom_module = pythoncom_module or loaded_pythoncom
        dispatch_factory = dispatch_factory or loaded_dispatch

    application: Any | None = None
    document: Any | None = None
    com_initialized = False
    processing_error: Exception | None = None
    cleanup_errors: list[str] = []
    temporary = _temporary_output(output)
    try:
        pythoncom_module.CoInitialize()
        com_initialized = True
        try:
            application = dispatch_factory(spec.program_id)
        except Exception as exc:
            raise LegacyOfficeError(
                f"{spec.program_id} の COM 登録を利用できません。Microsoft Office のインストールを確認してください。"
            ) from exc
        document = _open_document(application, source, spec)
        _save_document(document, temporary, spec)
        if not temporary.is_file() or temporary.stat().st_size == 0:
            raise LegacyOfficeError(f"Office が出力ファイルを作成しませんでした: {output}")
    except Exception as exc:
        processing_error = exc
        _remove_file(temporary)
        if isinstance(exc, LegacyOfficeError):
            raise
        raise LegacyOfficeError(
            f"{source.name} の変換に失敗しました: {exc}"
        ) from exc
    finally:
        if document is not None:
            try:
                _close_document(document, spec)
            except Exception as exc:  # pragma: no cover - COM 固有
                cleanup_errors.append(f"文書を閉じられません: {exc}")
        if application is not None:
            try:
                application.Quit()
            except Exception as exc:  # pragma: no cover - COM 固有
                cleanup_errors.append(f"Office を終了できません: {exc}")
        if com_initialized:
            try:
                pythoncom_module.CoUninitialize()
            except Exception as exc:  # pragma: no cover - COM 固有
                cleanup_errors.append(f"COM を解放できません: {exc}")
        if cleanup_errors and processing_error is None:
            _remove_file(temporary)
            raise LegacyOfficeError("; ".join(cleanup_errors))
    _commit_temporary(temporary, output)
    return output


def _batch_plan(input_dir: Path, output_dir: Path) -> list[tuple[Path, Path]]:
    source_root = input_dir.expanduser().resolve()
    if not source_root.is_dir():
        raise LegacyOfficeError(f"入力フォルダが見つかりません: {input_dir}")
    destination_root = _absolute_without_resolving(output_dir)
    _reject_link_ancestors(destination_root)
    if _lexists(destination_root):
        if _is_link_like(destination_root):
            raise LegacyOfficeError(
                f"出力フォルダがシンボリックリンクです: {destination_root}"
            )
        if not destination_root.is_dir():
            raise LegacyOfficeError(f"出力先がフォルダではありません: {destination_root}")
    candidates = sorted(
        path for path in source_root.rglob("*") if path.is_file() and path.suffix.lower() in CONVERSIONS
    )
    if not candidates:
        raise LegacyOfficeError(f".xls、.doc、.ppt が見つかりません: {source_root}")
    plan: list[tuple[Path, Path]] = []
    planned_outputs: dict[str, Path] = {}
    for source in candidates:
        spec = conversion_spec(source)
        relative = source.relative_to(source_root).with_suffix(spec.output_suffix)
        output = destination_root / relative
        _assert_new_output_path(output)
        output_key = str(output).casefold()
        if output_key in planned_outputs:
            raise LegacyOfficeError(
                "複数の入力が同じ出力先になります。バッチ処理を開始しません: "
                f"{planned_outputs[output_key]} / {output}"
            )
        planned_outputs[output_key] = output
        plan.append((source, output))
    return plan


def batch_convert(
    input_dir: Path,
    output_dir: Path,
    *,
    converter: Callable[[Path, Path], Path] | None = None,
) -> list[Path]:
    """入力フォルダを再帰走査し、相対構成を保って旧形式を一括変換する。"""
    plan = _batch_plan(input_dir, output_dir)
    convert_one = converter or (lambda source, output: convert_file(source, output))
    results: list[Path] = []
    attempted_outputs: list[Path] = []
    try:
        for source, output in plan:
            _assert_new_output_path(output)
            output.parent.mkdir(parents=True, exist_ok=True)
            _reject_link_ancestors(output)
            attempted_outputs.append(output)
            convert_one(source, output)
            if not output.is_file() or _is_link_like(output):
                raise LegacyOfficeError(f"変換後ファイルが作成されませんでした: {output}")
            results.append(output)
    except Exception as exc:
        rollback_errors: list[str] = []
        for attempted in reversed(attempted_outputs):
            try:
                _remove_file(attempted)
            except OSError as rollback_error:
                rollback_errors.append(f"{attempted}: {rollback_error}")
        if rollback_errors:
            raise LegacyOfficeError(
                "バッチ変換に失敗し、一部出力をロールバックできませんでした: "
                + "; ".join(rollback_errors)
            ) from exc
        raise
    return results


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Microsoft Office COM で旧 Office 形式を新形式へ変換します。"
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    convert = subparsers.add_parser("convert", help="単一ファイルを変換します。")
    convert.add_argument("input", type=Path)
    convert.add_argument("output", type=Path)

    batch = subparsers.add_parser("batch", help="フォルダ内の旧 Office ファイルを一括変換します。")
    batch.add_argument("input_dir", type=Path)
    batch.add_argument("output_dir", type=Path)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        _require_python_version()
        if args.command == "convert":
            results = [convert_file(args.input, args.output)]
        else:
            results = batch_convert(args.input_dir, args.output_dir)
    except LegacyOfficeError as exc:
        print(f"エラー: {exc}", file=sys.stderr)
        return 2
    except Exception as exc:  # pragma: no cover - COM 固有例外の最終境界
        print(f"エラー: 旧 Office 変換に失敗しました: {exc}", file=sys.stderr)
        return 1
    for result in results:
        print(f"完了: {result}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
