from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from pptx import Presentation


REPOSITORY_ROOT = Path(__file__).resolve().parents[1]
EXCEL_INSPECTOR = (
    REPOSITORY_ROOT
    / "analyze-excel-to-markdown"
    / "scripts"
    / "inspect_excel.py"
)
POWERPOINT_INSPECTOR = (
    REPOSITORY_ROOT
    / "analyze-powerpoint-to-markdown"
    / "scripts"
    / "inspect_powerpoint.py"
)
MARKDOWN_VALIDATOR = (
    REPOSITORY_ROOT
    / "write-vscode-markdown"
    / "scripts"
    / "validate_markdown.py"
)
PDF_CONVERTER = (
    REPOSITORY_ROOT
    / "write-vscode-markdown"
    / "scripts"
    / "markdown_to_pdf.mjs"
)


def run(command: list[str], *, environment: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        command,
        cwd=REPOSITORY_ROOT,
        env=environment,
        check=False,
        capture_output=True,
        text=True,
    )


def first_excel_cell(payload: dict, coordinate: str) -> dict:
    for sheet in payload["sheets"]:
        for row in sheet["rows"]:
            for cell in row["cells"]:
                if cell["coordinate"] == coordinate:
                    return cell
    raise AssertionError(f"Excelセルが見つかりません: {coordinate}")


class OfficeToMarkdownEndToEndTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary = tempfile.TemporaryDirectory()
        self.directory = Path(self.temporary.name)

    def tearDown(self) -> None:
        self.temporary.cleanup()

    def create_excel(self) -> Path:
        path = self.directory / "requirements.xlsx"
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "要件"
        sheet["A1"] = "REQ-001"
        sheet["B1"] = "監査ログを90日保持する"
        workbook.save(path)
        return path

    def create_powerpoint(self) -> Path:
        path = self.directory / "overview.pptx"
        presentation = Presentation()
        slide = presentation.slides.add_slide(presentation.slide_layouts[1])
        slide.shapes.title.text = "運用概要"
        slide.placeholders[1].text = "毎日バックアップを確認する"
        presentation.save(path)
        return path

    def inspect_excel(self, source: Path) -> dict:
        output = self.directory / "excel.inspection.json"
        completed = run(
            [
                sys.executable,
                str(EXCEL_INSPECTOR),
                str(source),
                "--output",
                str(output),
            ]
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        return json.loads(output.read_text(encoding="utf-8"))

    def inspect_powerpoint(self, source: Path) -> dict:
        output = self.directory / "powerpoint.inspection.json"
        completed = run(
            [
                sys.executable,
                str(POWERPOINT_INSPECTOR),
                str(source),
                "--output",
                str(output),
            ]
        )
        self.assertEqual(completed.returncode, 0, completed.stderr)
        return json.loads(output.read_text(encoding="utf-8"))

    def create_markdown(self, excel_payload: dict, powerpoint_payload: dict) -> Path:
        requirement_id = first_excel_cell(excel_payload, "A1")["value"]
        requirement = first_excel_cell(excel_payload, "B1")["value"]
        slide_title = powerpoint_payload["slides"][0]["title"]["text"]
        markdown = f"""# Office文書変換E2E

## 目次

1. [1. Excel解析](#1-excel解析)
2. [2. PowerPoint解析](#2-powerpoint解析)

## 1. Excel解析

| 要件ID | 内容 | 元セル |
|---|---|---|
| {requirement_id} | {requirement} | `'要件'!A1:B1` |

## 2. PowerPoint解析

スライドタイトルは「{slide_title}」です。`スライド 1`
"""
        path = self.directory / "office-analysis.md"
        path.write_text(markdown, encoding="utf-8", newline="\n")
        return path

    def test_excel_and_powerpoint_feed_valid_markdown(self) -> None:
        excel_payload = self.inspect_excel(self.create_excel())
        powerpoint_payload = self.inspect_powerpoint(self.create_powerpoint())
        markdown = self.create_markdown(excel_payload, powerpoint_payload)

        completed = run([sys.executable, str(MARKDOWN_VALIDATOR), str(markdown)])
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
        self.assertIn("REQ-001", markdown.read_text(encoding="utf-8"))
        self.assertIn("運用概要", markdown.read_text(encoding="utf-8"))

    @unittest.skipUnless(
        os.environ.get("RUN_PDF_E2E") == "1",
        "RUN_PDF_E2E=1のときだけPlaywright Chromiumを使うPDF試験を実行します",
    )
    def test_excel_and_powerpoint_feed_pdf(self) -> None:
        if shutil.which("node") is None:
            self.skipTest("Node.jsが見つかりません")

        excel_payload = self.inspect_excel(self.create_excel())
        powerpoint_payload = self.inspect_powerpoint(self.create_powerpoint())
        markdown = self.create_markdown(excel_payload, powerpoint_payload)
        pdf = self.directory / "office-analysis.pdf"

        completed = run(
            ["node", str(PDF_CONVERTER), str(markdown), str(pdf)],
            environment=os.environ.copy(),
        )
        self.assertEqual(completed.returncode, 0, completed.stdout + completed.stderr)
        self.assertTrue(pdf.read_bytes().startswith(b"%PDF-"))
        self.assertGreater(pdf.stat().st_size, 1_000)


if __name__ == "__main__":
    unittest.main()
