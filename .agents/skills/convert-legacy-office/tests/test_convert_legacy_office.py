from __future__ import annotations

import importlib.util
import os
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "convert_legacy_office.py"
SPEC = importlib.util.spec_from_file_location("legacy_office_under_test", SCRIPT)
assert SPEC and SPEC.loader
legacy = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = legacy
SPEC.loader.exec_module(legacy)


class FakePythonCom:
    def __init__(self) -> None:
        self.initialized = 0
        self.uninitialized = 0

    def CoInitialize(self) -> None:
        self.initialized += 1

    def CoUninitialize(self) -> None:
        self.uninitialized += 1


class FakeDocument:
    def __init__(self, fail: bool = False) -> None:
        self.fail = fail
        self.saved_path: Path | None = None
        self.file_format: int | None = None
        self.closed = False

    def _save(self, path: str, file_format: int) -> None:
        self.saved_path = Path(path)
        self.file_format = file_format
        self.saved_path.write_bytes(b"partial" if self.fail else b"converted")
        if self.fail:
            raise RuntimeError("save failed")

    def SaveAs(self, path: str, FileFormat: int) -> None:
        self._save(path, FileFormat)

    def SaveAs2(self, path: str, FileFormat: int) -> None:
        self._save(path, FileFormat)

    def Close(self, *args) -> None:
        self.closed = True


class FakeCollection:
    def __init__(self, document: FakeDocument) -> None:
        self.document = document
        self.opened = False

    def Open(self, *args, **kwargs) -> FakeDocument:
        self.opened = True
        return self.document


class FakeApplication:
    def __init__(self, program_id: str, document: FakeDocument) -> None:
        self.program_id = program_id
        self.Workbooks = FakeCollection(document)
        self.Documents = FakeCollection(document)
        self.Presentations = FakeCollection(document)
        self.quit_called = False
        self.Visible = None
        self.DisplayAlerts = None

    def Quit(self) -> None:
        self.quit_called = True


class LegacyOfficeTests(unittest.TestCase):
    def make_symlink(self, target: Path, link: Path, *, directory: bool = False) -> None:
        try:
            os.symlink(target, link, target_is_directory=directory)
        except (NotImplementedError, OSError) as exc:
            self.skipTest(f"シンボリックリンクを作成できません: {exc}")

    def test_conversion_mappings_and_cleanup_for_all_formats(self) -> None:
        cases = [
            ("book.xls", "book.xlsx", "Excel.Application", 51),
            ("design.doc", "design.docx", "Word.Application", 16),
            ("slides.ppt", "slides.pptx", "PowerPoint.Application", 24),
        ]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for input_name, output_name, expected_program, expected_format in cases:
                with self.subTest(input_name=input_name):
                    source = root / input_name
                    output = root / output_name
                    source.write_bytes(b"legacy")
                    document = FakeDocument()
                    application = FakeApplication(expected_program, document)
                    pythoncom = FakePythonCom()
                    programs: list[str] = []

                    def dispatch(program_id: str):
                        programs.append(program_id)
                        return application

                    legacy.convert_file(
                        source,
                        output,
                        platform="win32",
                        dispatch_factory=dispatch,
                        pythoncom_module=pythoncom,
                    )

                    self.assertEqual(programs, [expected_program])
                    self.assertEqual(document.file_format, expected_format)
                    self.assertTrue(document.closed)
                    self.assertTrue(application.quit_called)
                    self.assertEqual(pythoncom.initialized, 1)
                    self.assertEqual(pythoncom.uninitialized, 1)
                    self.assertEqual(output.read_bytes(), b"converted")

    def test_failure_removes_partial_output_and_closes_everything(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "design.doc"
            output = root / "design.docx"
            source.write_bytes(b"legacy")
            document = FakeDocument(fail=True)
            application = FakeApplication("Word.Application", document)
            pythoncom = FakePythonCom()

            with self.assertRaises(legacy.LegacyOfficeError):
                legacy.convert_file(
                    source,
                    output,
                    platform="win32",
                    dispatch_factory=lambda _: application,
                    pythoncom_module=pythoncom,
                )

            self.assertFalse(output.exists())
            self.assertTrue(document.closed)
            self.assertTrue(application.quit_called)
            self.assertEqual(pythoncom.uninitialized, 1)

    def test_non_windows_is_rejected_before_dispatch(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "book.xls"
            output = root / "book.xlsx"
            source.write_bytes(b"legacy")
            called = False

            def dispatch(_: str):
                nonlocal called
                called = True
                return object()

            with self.assertRaises(legacy.LegacyOfficeError):
                legacy.convert_file(
                    source,
                    output,
                    platform="linux",
                    dispatch_factory=dispatch,
                    pythoncom_module=FakePythonCom(),
                )
            self.assertFalse(called)
            self.assertFalse(output.exists())

    def test_broken_output_symlink_and_symlink_ancestor_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "book.xls"
            source.write_bytes(b"legacy")
            broken_output = root / "book.xlsx"
            self.make_symlink(root / "missing.xlsx", broken_output)
            called = False

            def dispatch(_: str):
                nonlocal called
                called = True
                return object()

            with self.assertRaises(legacy.LegacyOfficeError):
                legacy.convert_file(
                    source,
                    broken_output,
                    platform="win32",
                    dispatch_factory=dispatch,
                    pythoncom_module=FakePythonCom(),
                )
            self.assertFalse(called)
            self.assertTrue(os.path.lexists(broken_output))

            outside = root / "outside"
            outside.mkdir()
            linked_directory = root / "linked-output"
            self.make_symlink(outside, linked_directory, directory=True)
            with self.assertRaises(legacy.LegacyOfficeError):
                legacy.convert_file(
                    source,
                    linked_directory / "book.xlsx",
                    platform="win32",
                    dispatch_factory=dispatch,
                    pythoncom_module=FakePythonCom(),
                )
            self.assertFalse((outside / "book.xlsx").exists())

    def test_batch_preserves_relative_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            inputs = root / "input"
            outputs = root / "output"
            (inputs / "nested").mkdir(parents=True)
            (inputs / "book.xls").write_bytes(b"xls")
            (inputs / "nested" / "design.doc").write_bytes(b"doc")
            (inputs / "ignore.txt").write_text("skip", encoding="utf-8")

            def converter(source: Path, output: Path) -> Path:
                output.parent.mkdir(parents=True, exist_ok=True)
                output.write_text(source.name, encoding="utf-8")
                return output

            results = legacy.batch_convert(inputs, outputs, converter=converter)

            self.assertEqual(
                {path.relative_to(outputs).as_posix() for path in results},
                {"book.xlsx", "nested/design.docx"},
            )

    def test_existing_batch_output_stops_before_conversion(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            inputs = root / "input"
            outputs = root / "output"
            inputs.mkdir()
            outputs.mkdir()
            (inputs / "book.xls").write_bytes(b"xls")
            (outputs / "book.xlsx").write_bytes(b"keep")
            calls: list[Path] = []

            with self.assertRaises(legacy.LegacyOfficeError):
                legacy.batch_convert(
                    inputs,
                    outputs,
                    converter=lambda source, output: calls.append(output),
                )
            self.assertEqual(calls, [])
            self.assertEqual((outputs / "book.xlsx").read_bytes(), b"keep")

    def test_batch_failure_rolls_back_every_output_from_this_attempt(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            inputs = root / "input"
            outputs = root / "output"
            inputs.mkdir()
            first_input = inputs / "first.xls"
            second_input = inputs / "second.doc"
            first_input.write_bytes(b"first-original")
            second_input.write_bytes(b"second-original")
            call_count = 0

            def converter(source: Path, output: Path) -> Path:
                nonlocal call_count
                call_count += 1
                output.write_bytes(f"created-{source.name}".encode())
                if call_count == 2:
                    raise RuntimeError("second conversion failed")
                return output

            with self.assertRaisesRegex(RuntimeError, "second conversion failed"):
                legacy.batch_convert(inputs, outputs, converter=converter)

            self.assertFalse((outputs / "first.xlsx").exists())
            self.assertFalse((outputs / "second.docx").exists())
            self.assertEqual(first_input.read_bytes(), b"first-original")
            self.assertEqual(second_input.read_bytes(), b"second-original")


if __name__ == "__main__":
    unittest.main()
