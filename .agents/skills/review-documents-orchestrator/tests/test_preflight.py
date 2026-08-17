import importlib.util
import tempfile
import unittest
import zipfile
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).parents[1] / "scripts" / "preflight.py"
SPEC = importlib.util.spec_from_file_location("preflight", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


class PreflightTest(unittest.TestCase):
    def test_missing_root_fails(self):
        with tempfile.TemporaryDirectory() as temporary:
            missing_root = Path(temporary) / "missing"
            with (
                patch.object(MODULE, "_check_imports", return_value=([], [])),
                patch.object(MODULE, "_check_office_com", return_value=({}, [])),
            ):
                report = MODULE.run_checks(missing_root)
            self.assertFalse(report["ok"])
            self.assertTrue(any("リポジトリルート" in item for item in report["missing"]))

    def test_pdf_requires_pdf_and_ocr_dependencies(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for relative in ("input/checklists", "input/references", "input/targets"):
                (root / relative).mkdir(parents=True)
            (root / "input/targets/scan.pdf").write_bytes(b"pdf")
            with (
                patch.object(MODULE, "_check_imports", return_value=([], [])) as imports,
                patch.object(MODULE, "_check_tesseract", return_value=({}, [])) as ocr,
                patch.object(MODULE, "_check_office_com", return_value=({}, [])),
            ):
                MODULE.run_checks(root)
            required = imports.call_args.args[0]
            self.assertTrue(
                {"pypdf", "pdfplumber", "PIL", "pytesseract"} <= required
            )
            self.assertNotIn("reportlab", required)
            ocr.assert_called_once()

    def test_inventory_ignores_readme(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / "input/targets").mkdir(parents=True)
            (root / "input/targets/README.md").write_text("help", encoding="utf-8")
            files, extensions = MODULE.inventory_extensions(root)
            self.assertEqual([], files)
            self.assertEqual(set(), extensions)

    def test_report_does_not_persist_absolute_repository_root(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            with (
                patch.object(MODULE, "_check_imports", return_value=([], [])),
                patch.object(MODULE, "_check_office_com", return_value=({}, [])),
            ):
                report = MODULE.run_checks(root)
            self.assertEqual(".", report["root"])
            self.assertNotIn(str(root), str(report))

    def test_report_rejects_input_symlink(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "repo"
            target_directory = root / "input/targets"
            target_directory.mkdir(parents=True)
            external = Path(temporary) / "outside.docx"
            external.write_bytes(b"outside")
            try:
                (target_directory / "outside.docx").symlink_to(external)
            except OSError as exc:
                self.skipTest(f"symlinkを作成できません: {exc}")
            with (
                patch.object(MODULE, "_check_imports", return_value=([], [])),
                patch.object(MODULE, "_check_office_com", return_value=({}, [])),
            ):
                report = MODULE.run_checks(root)
            self.assertFalse(report["ok"])
            self.assertTrue(any("symlink" in item for item in report["missing"]))

    def test_inventory_rejects_junction_like_input_ancestor(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = (Path(temporary) / "repo").resolve()
            junction = root / "input"
            (junction / "checklists").mkdir(parents=True)
            with patch.object(
                type(root),
                "is_junction",
                autospec=True,
                side_effect=lambda candidate: candidate == junction,
            ):
                with self.assertRaisesRegex(ValueError, "symlink/junction"):
                    MODULE.inventory_extensions(root)

    def test_office_embedded_image_requires_ocr(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            for relative in ("input/checklists", "input/references", "input/targets"):
                (root / relative).mkdir(parents=True)
            presentation = root / "input/targets/design.pptx"
            with zipfile.ZipFile(presentation, "w") as archive:
                archive.writestr("ppt/media/image1.png", b"png")
            with (
                patch.object(MODULE, "_check_imports", return_value=([], [])) as imports,
                patch.object(MODULE, "_check_tesseract", return_value=({}, [])) as ocr,
                patch.object(MODULE, "_check_office_com", return_value=({}, [])),
            ):
                report = MODULE.run_checks(root)
            self.assertTrue(report["office_embedded_images"])
            self.assertTrue({"PIL", "pytesseract"} <= imports.call_args.args[0])
            ocr.assert_called_once()


if __name__ == "__main__":
    unittest.main()
