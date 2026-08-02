from __future__ import annotations

import importlib.util
import sys
import tempfile
import unittest
from pathlib import Path


SCRIPT = Path(__file__).parents[1] / "scripts" / "validate_markdown.py"
SPEC = importlib.util.spec_from_file_location("write_vscode_validate_markdown", SCRIPT)
assert SPEC and SPEC.loader
validator = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = validator
SPEC.loader.exec_module(validator)


def document(body: str) -> str:
    return (
        "# テスト文書\n\n"
        "## 目次\n\n"
        "1. [1. 本文](#1-本文)\n\n"
        "## 1. 本文\n\n"
        f"{body.rstrip()}\n"
    )


def codes(path: Path) -> set[str]:
    return {item.code for item in validator.validate_file(path)}


class HeadingSlugTests(unittest.TestCase):
    def test_unique_slugs_match_github_style_collision_handling(self) -> None:
        lines = ["# Foo", "## Foo", "## Foo-1", "## Foo"]
        scan = validator.scan_blocks(Path("headings.md"), lines)
        headings, _ = validator.collect_headings(Path("headings.md"), scan)
        self.assertEqual(
            [heading.slug for heading in headings],
            ["foo", "foo-1", "foo-1-1", "foo-2"],
        )

    def test_markdown_escapes_and_entities_are_removed_consistently(self) -> None:
        self.assertEqual(validator.heading_slug(r"A \&amp; B"), "a--b")
        self.assertEqual(validator.heading_slug("&copy;"), "©")
        self.assertEqual(validator.heading_slug("&copy"), "copy")


class LinkValidationTests(unittest.TestCase):
    def test_checks_non_toc_anchor_and_missing_relative_target(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "doc.md"
            markdown.write_text(
                document("[bad anchor](#missing)\n\n[bad file](./missing.md)"),
                encoding="utf-8",
            )
            self.assertTrue({"LINK001", "LINK002"}.issubset(codes(markdown)))

    def test_checks_anchor_in_linked_markdown(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            linked = root / "linked.md"
            linked.write_text("# Linked\n\n## 1. Section\n", encoding="utf-8")
            markdown = root / "doc.md"
            markdown.write_text(
                document("[good](linked.md#1-section)\n\n[bad](linked.md#missing)"),
                encoding="utf-8",
            )
            diagnostics = validator.validate_file(markdown)
            link3 = [item for item in diagnostics if item.code == "LINK003"]
            self.assertEqual(len(link3), 1)
            self.assertIn("#missing", link3[0].message)


class PipeTableTests(unittest.TestCase):
    def test_reports_header_separator_and_body_column_mismatches(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "doc.md"
            markdown.write_text(
                document("| A | B |\n|---|\n| 1 | 2 | 3 |"), encoding="utf-8"
            )
            self.assertTrue({"TABLE001", "TABLE002"}.issubset(codes(markdown)))

    def test_escaped_pipe_does_not_add_a_column(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "doc.md"
            markdown.write_text(
                document("| A | B |\n|---|---|\n| left \\| right | value |"),
                encoding="utf-8",
            )
            self.assertFalse({"TABLE001", "TABLE002"} & codes(markdown))


class ImageAndSvgTests(unittest.TestCase):
    def test_query_and_fragment_are_not_part_of_image_filesystem_path(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "asset.png").write_bytes(b"png")
            markdown = root / "doc.md"
            markdown.write_text(
                document("![asset](asset.png?version=1#preview)"), encoding="utf-8"
            )
            self.assertFalse({"IMG003", "IMG005", "IMG006"} & codes(markdown))

    def test_parent_traversal_is_rejected_even_when_asset_exists(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "outside.png").write_bytes(b"png")
            docs = root / "docs"
            docs.mkdir()
            markdown = docs / "doc.md"
            markdown.write_text(
                document("![outside](../outside.png?version=1#preview)"), encoding="utf-8"
            )
            self.assertIn("IMG005", codes(markdown))

    def test_svg_rejects_file_relative_and_css_external_references(self) -> None:
        samples = {
            "file.svg": '<svg><image href="file:///tmp/image.png"/></svg>',
            "relative.svg": '<svg><use href="other.svg#shape"/></svg>',
            "css.svg": "<svg><style>.x { fill: url(other.svg#paint); }</style></svg>",
        }
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name, source in samples.items():
                with self.subTest(name=name):
                    (root / name).write_text(source, encoding="utf-8")
                    markdown = root / "doc.md"
                    markdown.write_text(document(f"![svg]({name})"), encoding="utf-8")
                    self.assertIn("SVG001", codes(markdown))

    def test_svg_allows_document_local_fragment_references(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "safe.svg").write_text(
                '<svg><defs><linearGradient id="g"/></defs>'
                '<rect id="shape" style="fill:url(#g)"/>'
                '<use href="#shape"/></svg>',
                encoding="utf-8",
            )
            markdown = root / "doc.md"
            markdown.write_text(document("![safe](safe.svg)"), encoding="utf-8")
            self.assertNotIn("SVG001", codes(markdown))


class MermaidCaseTests(unittest.TestCase):
    def test_language_identifier_must_be_lowercase(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "doc.md"
            markdown.write_text(
                document("```Mermaid\nflowchart LR\n```"), encoding="utf-8"
            )
            self.assertIn("MER010", codes(markdown))

    def test_diagram_declaration_is_case_sensitive(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            markdown = Path(directory) / "doc.md"
            markdown.write_text(
                document(
                    "```mermaid\n"
                    "FLOWCHART LR\n"
                    "  accTitle: Test\n"
                    "  accDescr: Test description\n"
                    "```"
                ),
                encoding="utf-8",
            )
            self.assertIn("MER002", codes(markdown))


if __name__ == "__main__":
    unittest.main()
