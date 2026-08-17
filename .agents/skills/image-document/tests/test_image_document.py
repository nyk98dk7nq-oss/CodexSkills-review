from __future__ import annotations

import importlib.util
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).resolve().parents[1] / "scripts" / "image_document.py"
SPEC = importlib.util.spec_from_file_location("image_document_under_test", SCRIPT)
assert SPEC and SPEC.loader
image_document = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(image_document)


try:
    from PIL import Image
except ImportError:  # pragma: no cover - テスト環境依存
    Image = None


@unittest.skipIf(Image is None, "Pillow が必要です")
class ImageDocumentTests(unittest.TestCase):
    def make_symlink(self, target: Path, link: Path, *, directory: bool = False) -> None:
        try:
            os.symlink(target, link, target_is_directory=directory)
        except (NotImplementedError, OSError) as exc:
            self.skipTest(f"シンボリックリンクを作成できません: {exc}")

    def make_multiframe_tiff(self, path: Path) -> None:
        first = Image.new("RGB", (40, 30), "white")
        second = Image.new("RGB", (40, 30), "navy")
        first.save(path, format="TIFF", save_all=True, append_images=[second])

    def test_to_markdown_reads_all_frames_and_exports_frames(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "scan.tiff"
            source.parent.mkdir(parents=True)
            self.make_multiframe_tiff(source)
            output = root / "work" / "markdown" / "scan.md"
            images = root / "work" / "images"

            image_document.image_to_markdown(
                source,
                output,
                role="target",
                repo_root=root,
                images_dir=images,
            )

            markdown = output.read_text(encoding="utf-8")
            self.assertIn('source_path: "input/targets/scan.tiff"', markdown)
            self.assertIn('source_format: "tiff"', markdown)
            self.assertIn("image_width_px: 40", markdown)
            self.assertIn("image_height_px: 30", markdown)
            self.assertIn("frame_count: 2", markdown)
            self.assertIn("ocr_executed: false", markdown)
            self.assertIn("ocr_languages: []", markdown)
            self.assertIn("## フレーム 2", markdown)
            self.assertIn(
                "![scan.tiff フレーム 1](../images/scan-frame-0001.png)",
                markdown,
            )
            self.assertIn(
                "![scan.tiff フレーム 2](../images/scan-frame-0002.png)",
                markdown,
            )
            self.assertTrue((images / "scan-frame-0001.png").is_file())
            self.assertTrue((images / "scan-frame-0002.png").is_file())

    def test_frontmatter_preserves_original_image_extension(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            target_dir = root / "input" / "targets"
            target_dir.mkdir(parents=True)
            for suffix, image_format in (("jpg", "JPEG"), ("jpeg", "JPEG"), ("tif", "TIFF")):
                with self.subTest(suffix=suffix):
                    source = target_dir / f"sample.{suffix}"
                    output = root / "work" / f"sample-{suffix}.md"
                    Image.new("RGB", (12, 8), "white").save(source, format=image_format)

                    image_document.image_to_markdown(
                        source, output, role="target", repo_root=root
                    )

                    markdown = output.read_text(encoding="utf-8")
                    self.assertIn(f'source_format: "{suffix}"', markdown)
                    self.assertIn(
                        f"![sample.{suffix} フレーム 1](../input/targets/sample.{suffix})",
                        markdown,
                    )

    def test_markdown_link_normalizes_windows_separators(self) -> None:
        with patch.object(
            image_document.os.path,
            "relpath",
            return_value=r"..\images\frame 1.png",
        ):
            link = image_document._relative_markdown_link(
                Path("frame.png"), Path("output.md")
            )
        self.assertEqual(link, "../images/frame%201.png")

    def test_ocr_frontmatter_uses_readme_schema_and_language_array(self) -> None:
        class FakeOutput:
            DICT = object()

        class FakeTesseract:
            Output = FakeOutput

            @staticmethod
            def image_to_data(*args, **kwargs):
                return {
                    "text": ["設計"],
                    "conf": ["91.25"],
                    "left": [5],
                    "top": [6],
                    "width": [20],
                    "height": [8],
                }

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "diagram.png"
            source.parent.mkdir(parents=True)
            Image.new("RGB", (80, 40), "white").save(source)
            output = root / "work" / "diagram.md"
            with patch.object(
                image_document, "_load_pytesseract", return_value=FakeTesseract()
            ):
                image_document.image_to_markdown(
                    source,
                    output,
                    role="target",
                    repo_root=root,
                    ocr=True,
                    lang="jpn+eng+jpn_vert",
                )

            markdown = output.read_text(encoding="utf-8")
            self.assertIn("image_width_px: 80", markdown)
            self.assertIn("image_height_px: 40", markdown)
            self.assertIn("ocr_executed: true", markdown)
            self.assertIn(
                'ocr_languages:\n  - "jpn"\n  - "eng"\n  - "jpn_vert"',
                markdown,
            )

    def test_edit_crops_resizes_and_grayscales_all_frames(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.tiff"
            output = root / "edited.tiff"
            operations = root / "operations.json"
            self.make_multiframe_tiff(source)
            operations.write_text(
                json.dumps(
                    {
                        "operations": [
                            {"op": "crop", "box": [0, 0, 20, 20]},
                            {"op": "resize", "width": 10},
                            {"op": "grayscale"},
                        ]
                    }
                ),
                encoding="utf-8",
            )

            image_document.edit_image(source, output, operations)

            with Image.open(output) as edited:
                self.assertEqual(edited.n_frames, 2)
                for index in range(2):
                    edited.seek(index)
                    self.assertEqual(edited.size, (10, 10))
                    self.assertEqual(edited.mode, "L")

    def test_existing_output_is_rejected_without_change(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.png"
            output = root / "result.md"
            Image.new("RGB", (10, 10), "white").save(source)
            output.write_text("keep", encoding="utf-8")

            with self.assertRaises(image_document.ImageDocumentError):
                image_document.image_to_markdown(
                    source, output, role="target", repo_root=root
                )
            self.assertEqual(output.read_text(encoding="utf-8"), "keep")

    def test_broken_output_symlink_and_symlink_ancestor_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "source.png"
            source.parent.mkdir(parents=True)
            Image.new("RGB", (10, 10), "white").save(source)

            broken_output = root / "broken.md"
            self.make_symlink(root / "missing.md", broken_output)
            with self.assertRaises(image_document.ImageDocumentError):
                image_document.image_to_markdown(
                    source, broken_output, role="target", repo_root=root
                )
            self.assertTrue(os.path.lexists(broken_output))

            outside = root / "outside"
            outside.mkdir()
            linked_directory = root / "linked-output"
            self.make_symlink(outside, linked_directory, directory=True)
            with self.assertRaises(image_document.ImageDocumentError):
                image_document.image_to_markdown(
                    source,
                    linked_directory / "escaped.md",
                    role="target",
                    repo_root=root,
                )
            self.assertFalse((outside / "escaped.md").exists())

    def test_images_directory_symlink_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "source.png"
            source.parent.mkdir(parents=True)
            Image.new("RGB", (10, 10), "white").save(source)
            outside = root / "outside-images"
            outside.mkdir()
            images_link = root / "images-link"
            self.make_symlink(outside, images_link, directory=True)

            with self.assertRaises(image_document.ImageDocumentError):
                image_document.image_to_markdown(
                    source,
                    root / "work" / "source.md",
                    role="target",
                    repo_root=root,
                    images_dir=images_link,
                )
            self.assertEqual(list(outside.iterdir()), [])

    def test_all_frame_conflicts_are_checked_before_first_write(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "input" / "targets" / "scan.tiff"
            source.parent.mkdir(parents=True)
            self.make_multiframe_tiff(source)
            images = root / "work" / "images"
            images.mkdir(parents=True)
            second_frame = images / "scan-frame-0002.png"
            self.make_symlink(root / "missing-frame.png", second_frame)

            with self.assertRaises(image_document.ImageDocumentError):
                image_document.image_to_markdown(
                    source,
                    root / "work" / "scan.md",
                    role="target",
                    repo_root=root,
                    images_dir=images,
                )

            self.assertFalse((images / "scan-frame-0001.png").exists())
            self.assertTrue(os.path.lexists(second_frame))

    def test_invalid_multiframe_output_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "source.tiff"
            output = root / "edited.jpg"
            operations = root / "operations.json"
            self.make_multiframe_tiff(source)
            operations.write_text('[{"op": "grayscale"}]', encoding="utf-8")

            with self.assertRaises(image_document.ImageDocumentError):
                image_document.edit_image(source, output, operations)
            self.assertFalse(output.exists())

    def test_ocr_rows_include_coordinates_and_confidence(self) -> None:
        class FakeOutput:
            DICT = object()

        class FakeTesseract:
            Output = FakeOutput

            @staticmethod
            def image_to_data(*args, **kwargs):
                return {
                    "text": ["設計"],
                    "conf": ["91.25"],
                    "left": [5],
                    "top": [6],
                    "width": [20],
                    "height": [8],
                }

        rows = image_document._ocr_rows(object(), FakeTesseract(), "jpn+eng")
        self.assertEqual(rows[0]["text"], "設計")
        self.assertEqual(rows[0]["confidence"], 91.25)
        self.assertEqual(rows[0]["x"], 5)
        self.assertEqual(rows[0]["height"], 8)


if __name__ == "__main__":
    unittest.main()
