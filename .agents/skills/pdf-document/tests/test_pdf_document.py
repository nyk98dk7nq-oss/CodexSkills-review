from __future__ import annotations

import importlib.util
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "pdf_document.py"
SPEC = importlib.util.spec_from_file_location("pdf_document_under_test", SCRIPT)
assert SPEC and SPEC.loader
pdf_document = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(pdf_document)


try:
    from pypdf import PdfReader
    from PIL import Image
    from reportlab.pdfgen import canvas
except ImportError:  # pragma: no cover - テスト環境依存
    PdfReader = None
    Image = None
    canvas = None


@unittest.skipIf(
    PdfReader is None or canvas is None or Image is None,
    "pypdf、Pillow、reportlab が必要です",
)
class PdfDocumentTests(unittest.TestCase):
    def make_symlink(self, target: Path, link: Path, *, directory: bool = False) -> None:
        try:
            os.symlink(target, link, target_is_directory=directory)
        except (NotImplementedError, OSError) as exc:
            self.skipTest(f"シンボリックリンクを作成できません: {exc}")

    def make_pdf(self, path: Path, texts: list[str]) -> None:
        document = canvas.Canvas(str(path))
        for text in texts:
            document.drawString(72, 720, text)
            document.showPage()
        document.save()

    def test_to_markdown_extracts_pages_text_and_coordinates(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "design.pdf"
            source.parent.mkdir(parents=True)
            self.make_pdf(source, ["First page requirement", "Second page evidence"])
            output = root / "work" / "markdown" / "design.md"
            images = root / "work" / "images"

            result = pdf_document.pdf_to_markdown(
                source,
                output,
                role="target",
                repo_root=root,
                images_dir=images,
            )

            self.assertEqual(result, output)
            markdown = output.read_text(encoding="utf-8")
            self.assertIn('source_path: "input/targets/design.pdf"', markdown)
            self.assertIn('source_name: "design.pdf"', markdown)
            self.assertIn('document_role: "target"', markdown)
            self.assertIn('converter_skill: "pdf-document"', markdown)
            self.assertIn("## ページ 1", markdown)
            self.assertIn("First page requirement", markdown)
            self.assertIn("p.1 x=", markdown)
            self.assertIn("## ページ 2", markdown)
            self.assertTrue((images / "design-page-0001.png").is_file())
            self.assertTrue((images / "design-page-0002.png").is_file())

    def test_ocr_runs_for_page_with_text_and_embedded_image(self) -> None:
        class FakeOutput:
            DICT = object()

        class FakeTesseract:
            Output = FakeOutput

            @staticmethod
            def image_to_data(*args, **kwargs):
                return {
                    "text": ["画像内文字"],
                    "conf": ["96.0"],
                    "left": [15],
                    "top": [25],
                    "width": [60],
                    "height": [12],
                }

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "mixed.pdf"
            source.parent.mkdir(parents=True)
            embedded = root / "embedded.png"
            Image.new("RGB", (100, 40), "white").save(embedded)
            document = canvas.Canvas(str(source))
            document.drawString(
                72,
                720,
                "This page has enough ordinary text to exceed the sparse threshold.",
            )
            document.drawImage(str(embedded), 72, 600, width=100, height=40)
            document.showPage()
            document.save()
            output = root / "work" / "markdown" / "mixed.md"

            with patch.object(
                pdf_document, "_load_pytesseract", return_value=FakeTesseract()
            ):
                pdf_document.pdf_to_markdown(
                    source,
                    output,
                    role="target",
                    repo_root=root,
                    ocr=True,
                    lang="jpn+eng+jpn_vert",
                )

            markdown = output.read_text(encoding="utf-8")
            self.assertIn("### PDF 内画像領域", markdown)
            self.assertIn("画像内文字", markdown)
            self.assertIn("96.00", markdown)
            self.assertNotIn("テキストレイヤーが十分にあるため OCR を省略", markdown)

    def test_edit_reorders_rotates_and_updates_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.pdf"
            output = root / "edited.pdf"
            operations = root / "operations.json"
            self.make_pdf(source, ["Page A", "Page B"])
            operations.write_text(
                json.dumps(
                    {
                        "operations": [
                            {"op": "reorder", "pages": [2, 1]},
                            {"op": "rotate", "pages": [1], "degrees": 90},
                            {"op": "metadata", "values": {"Title": "Reviewed"}},
                        ]
                    }
                ),
                encoding="utf-8",
            )

            pdf_document.edit_pdf(source, output, operations)

            edited = PdfReader(str(output))
            self.assertEqual(len(edited.pages), 2)
            self.assertIn("Page B", edited.pages[0].extract_text())
            self.assertEqual(edited.pages[0].rotation, 90)
            self.assertEqual(edited.metadata.title, "Reviewed")

    def test_edit_merges_selected_page_and_deletes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.pdf"
            appendix = root / "appendix.pdf"
            output = root / "combined.pdf"
            operations = root / "operations.json"
            self.make_pdf(source, ["Main 1", "Main 2"])
            self.make_pdf(appendix, ["Appendix 1", "Appendix 2"])
            operations.write_text(
                json.dumps(
                    [
                        {"op": "delete", "pages": [2]},
                        {
                            "op": "merge",
                            "file": "appendix.pdf",
                            "pages": [2],
                            "position": "end",
                        },
                    ]
                ),
                encoding="utf-8",
            )

            pdf_document.edit_pdf(source, output, operations)

            combined = PdfReader(str(output))
            self.assertEqual(len(combined.pages), 2)
            self.assertIn("Appendix 2", combined.pages[1].extract_text())

    def test_existing_output_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.pdf"
            output = root / "result.md"
            self.make_pdf(source, ["Text"])
            output.write_text("keep", encoding="utf-8")

            with self.assertRaises(pdf_document.PdfDocumentError):
                pdf_document.pdf_to_markdown(
                    source, output, role="target", repo_root=root
                )
            self.assertEqual(output.read_text(encoding="utf-8"), "keep")

    def test_broken_output_symlink_and_symlink_ancestor_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "source.pdf"
            source.parent.mkdir(parents=True)
            self.make_pdf(source, ["Text"])

            broken_output = root / "broken.md"
            self.make_symlink(root / "missing.md", broken_output)
            with self.assertRaises(pdf_document.PdfDocumentError):
                pdf_document.pdf_to_markdown(
                    source, broken_output, role="target", repo_root=root
                )
            self.assertTrue(os.path.lexists(broken_output))

            outside = root / "outside"
            outside.mkdir()
            linked_directory = root / "linked-output"
            self.make_symlink(outside, linked_directory, directory=True)
            with self.assertRaises(pdf_document.PdfDocumentError):
                pdf_document.pdf_to_markdown(
                    source,
                    linked_directory / "escaped.md",
                    role="target",
                    repo_root=root,
                )
            self.assertFalse((outside / "escaped.md").exists())

    def test_images_directory_symlink_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "source.pdf"
            source.parent.mkdir(parents=True)
            self.make_pdf(source, ["Text"])
            outside = root / "outside-images"
            outside.mkdir()
            images_link = root / "images-link"
            self.make_symlink(outside, images_link, directory=True)

            with self.assertRaises(pdf_document.PdfDocumentError):
                pdf_document.pdf_to_markdown(
                    source,
                    root / "work" / "source.md",
                    role="target",
                    repo_root=root,
                    images_dir=images_link,
                )
            self.assertEqual(list(outside.iterdir()), [])

    def test_all_page_image_conflicts_are_checked_before_first_write(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "source.pdf"
            source.parent.mkdir(parents=True)
            self.make_pdf(source, ["Page one", "Page two"])
            images = root / "work" / "images"
            images.mkdir(parents=True)
            second_page = images / "source-page-0002.png"
            self.make_symlink(root / "missing-page.png", second_page)

            with self.assertRaises(pdf_document.PdfDocumentError):
                pdf_document.pdf_to_markdown(
                    source,
                    root / "work" / "source.md",
                    role="target",
                    repo_root=root,
                    images_dir=images,
                )

            self.assertFalse((images / "source-page-0001.png").exists())
            self.assertTrue(os.path.lexists(second_page))
            self.assertFalse((root / "work" / "source.md").exists())

    def test_ocr_rows_include_pixel_and_pdf_coordinates(self) -> None:
        class FakeOutput:
            DICT = object()

        class FakeTesseract:
            Output = FakeOutput

            @staticmethod
            def image_to_data(*args, **kwargs):
                return {
                    "text": ["要件"],
                    "conf": ["88.5"],
                    "left": [10],
                    "top": [20],
                    "width": [30],
                    "height": [40],
                }

        class FakeImage:
            size = (200, 400)

        rows = pdf_document._ocr_rows(FakeImage(), FakeTesseract(), "jpn+eng", 100, 200)
        self.assertEqual(rows[0]["text"], "要件")
        self.assertEqual(rows[0]["pdf_x0"], 5)
        self.assertEqual(rows[0]["pdf_bottom"], 30)


if __name__ == "__main__":
    unittest.main()
