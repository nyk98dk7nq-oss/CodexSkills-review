import importlib.util
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).parents[1] / "scripts" / "review_markdown_documents.py"
SPEC = importlib.util.spec_from_file_location("review_markdown_documents", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


def valid_data(result="不適合"):
    improvement = "エラーコードと発生条件を表に追加する。" if result == "不適合" else ""
    return {
        "run_id": "202601021530",
        "generated_at": "2026-01-02T15:30:00+09:00",
        "checklists": [
            {
                "path": "input/checklists/check.xlsx",
                "header_row": 1,
                "markdown": "check.md",
            }
        ],
        "references": [{"path": "input/references/ref.pdf", "markdown": "ref.md"}],
        "targets": [{"path": "input/targets/design.docx", "markdown": "target.md"}],
        "checklist_items": [
            {
                "checklist_file": "input/checklists/check.xlsx",
                "sheet": "Sheet1",
                "row": 2,
                "check_item": "エラー定義があるか",
            }
        ],
        "items": [
            {
                "checklist_file": "input/checklists/check.xlsx",
                "sheet": "Sheet1",
                "row": 2,
                "check_item": "エラー定義があるか",
                "target_file": "input/targets/design.docx",
                "result": result,
                "comment": "target.md:L10-L12にエラー定義がない。",
                "evidence": ["target.md:L10-L12"],
                "improvement": improvement,
            }
        ],
    }


class ReviewMarkdownDocumentsTest(unittest.TestCase):
    def _write_prepare_fixture(self, root):
        checklist_markdown = (
            root / "work/markdown/202601021530/checklists/check.xlsx.md"
        )
        target_markdown = root / "work/markdown/202601021530/targets/design.docx.md"
        checklist_markdown.parent.mkdir(parents=True)
        target_markdown.parent.mkdir(parents=True)
        checklist_markdown.write_text("# checklist\n", encoding="utf-8")
        target_markdown.write_text("# target\n", encoding="utf-8")
        (root / "input/checklists").mkdir(parents=True)
        (root / "input/targets").mkdir(parents=True)
        (root / "input/checklists/check.xlsx").write_bytes(b"xlsx")
        (root / "input/targets/design.docx").write_bytes(b"docx")
        manifest = {
            "root": ".",
            "run_id": "202601021530",
            "generated_at": "2026-01-02T15:30:00+09:00",
            "checklists": [
                {
                    "path": "input/checklists/check.xlsx",
                    "original_path": "input/checklists/check.xlsx",
                    "markdown": "work/markdown/202601021530/checklists/check.xlsx.md",
                    "header_row": 1,
                }
            ],
            "references": [],
            "targets": [
                {
                    "path": "input/targets/design.docx",
                    "original_path": "input/targets/design.docx",
                    "markdown": "work/markdown/202601021530/targets/design.docx.md",
                }
            ],
            "checklist_items": [
                {
                    "checklist_file": "input/checklists/check.xlsx",
                    "sheet": "Sheet1",
                    "row": 2,
                    "check_item": "項目A",
                }
            ],
        }
        manifest_path = root / "work/review-runs/202601021530/manifest.json"
        manifest_path.parent.mkdir(parents=True)
        manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
        return manifest_path

    def test_validate_requires_complete_cross_product(self):
        data = valid_data()
        data["targets"].append(
            {"path": "input/targets/other.docx", "markdown": "other.md"}
        )
        errors = MODULE.validate_results(data)
        self.assertTrue(any("不足している組合せ" in error for error in errors))

    def test_repo_path_rejects_junction_like_ancestor(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            junction = root / "work"
            junction.mkdir()
            output = junction / "review-runs/202601021530/review_bundle.json"
            with patch.object(
                type(root),
                "is_junction",
                autospec=True,
                side_effect=lambda candidate: candidate == junction,
            ):
                with self.assertRaisesRegex(
                    MODULE.ReviewDataError, "symlink/junction"
                ):
                    MODULE._validated_repo_path(root, output, "出力先")

    def test_validate_rejects_unknown_result_and_missing_improvement(self):
        data = valid_data()
        data["items"][0]["result"] = "OK"
        errors = MODULE.validate_results(data)
        self.assertTrue(any("resultは" in error for error in errors))

        data = valid_data()
        data["items"][0]["improvement"] = ""
        errors = MODULE.validate_results(data)
        self.assertTrue(any("具体的なimprovement" in error for error in errors))

    def test_summary_focuses_on_nonconformity(self):
        summary = MODULE.render_summary(valid_data())
        self.assertIn("## 不適合項目と改善案", summary)
        self.assertIn("エラーコードと発生条件を表に追加する。", summary)
        self.assertIn("target.md:L10-L12", summary)
        self.assertIn("| 不適合 | 1 |", summary)

    def test_summary_says_no_nonconformity(self):
        summary = MODULE.render_summary(valid_data("適合"))
        self.assertIn("不適合なし", summary)

    def test_prepare_creates_every_combination(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            check_markdown = root / "work/markdown/202601021530/checklists/check.xlsx.md"
            target_markdown = root / "work/markdown/202601021530/targets/target.md"
            check_markdown.parent.mkdir(parents=True)
            target_markdown.parent.mkdir(parents=True)
            check_markdown.write_text("# checklist\n", encoding="utf-8")
            target_markdown.write_text("# target\n", encoding="utf-8")
            manifest = {
                "root": ".",
                "run_id": "202601021530",
                "generated_at": "2026-01-02T15:30:00+09:00",
                "checklists": [
                    {
                        "path": "input/checklists/check.xlsx",
                        "original_path": "input/checklists/check.xlsx",
                        "markdown": "work/markdown/202601021530/checklists/check.xlsx.md",
                        "header_row": 1,
                    }
                ],
                "references": [],
                "targets": [
                    {
                        "path": "input/targets/a.docx",
                        "original_path": "input/targets/a.docx",
                        "markdown": "work/markdown/202601021530/targets/target.md",
                    },
                    {
                        "path": "input/targets/b.docx",
                        "original_path": "input/targets/b.docx",
                        "markdown": "work/markdown/202601021530/targets/target.md",
                    },
                ],
                "checklist_items": [
                    {
                        "checklist_file": "input/checklists/check.xlsx",
                        "sheet": "Sheet1",
                        "row": 2,
                        "check_item": "項目A",
                    }
                ],
            }
            manifest_path = root / "work/review-runs/202601021530/manifest.json"
            manifest_path.parent.mkdir(parents=True)
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            bundle = MODULE.prepare_bundle(manifest_path)
            self.assertEqual(2, len(bundle["items"]))
            self.assertEqual(3, len(bundle["materials"]))
            self.assertNotIn(str(root), json.dumps(bundle, ensure_ascii=False))

    def test_prepare_rejects_markdown_outside_repository(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            manifest = {
                "root": ".",
                "run_id": "202601021530",
                "generated_at": "2026-01-02T15:30:00+09:00",
                "checklists": [
                    {
                        "path": "input/checklists/check.xlsx",
                        "original_path": "input/checklists/check.xlsx",
                        "markdown": "../outside.md",
                        "header_row": 1,
                    }
                ],
                "references": [],
                "targets": [
                    {
                        "path": "input/targets/design.docx",
                        "original_path": "input/targets/design.docx",
                        "markdown": "work/markdown/202601021530/targets/target.md",
                    }
                ],
                "checklist_items": [
                    {
                        "checklist_file": "input/checklists/check.xlsx",
                        "sheet": "Sheet1",
                        "row": 2,
                        "check_item": "項目A",
                    }
                ],
            }
            target_markdown = root / "work/markdown/202601021530/targets/target.md"
            target_markdown.parent.mkdir(parents=True)
            target_markdown.write_text("# target\n", encoding="utf-8")
            manifest_path = root / "work/review-runs/202601021530/manifest.json"
            manifest_path.parent.mkdir(parents=True)
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            with self.assertRaisesRegex(MODULE.ReviewDataError, "root相対"):
                MODULE.prepare_bundle(manifest_path)

    def test_prepare_cli_never_overwrites_existing_or_input_file(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            manifest_path = self._write_prepare_fixture(root)
            bundle = manifest_path.parent / "review_bundle.json"
            bundle.write_text("keep", encoding="utf-8")
            code = MODULE.main(
                [
                    "prepare",
                    "--manifest",
                    str(manifest_path),
                    "--output",
                    str(bundle),
                ]
            )
            self.assertEqual(1, code)
            self.assertEqual("keep", bundle.read_text(encoding="utf-8"))

            input_path = root / "input/checklists/check.xlsx"
            code = MODULE.main(
                [
                    "prepare",
                    "--manifest",
                    str(manifest_path),
                    "--output",
                    str(input_path),
                ]
            )
            self.assertEqual(1, code)
            self.assertEqual(b"xlsx", input_path.read_bytes())

    def test_prepare_cli_writes_new_bundle_atomically(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            manifest_path = self._write_prepare_fixture(root)
            bundle = manifest_path.parent / "review_bundle.json"
            code = MODULE.main(
                [
                    "prepare",
                    "--manifest",
                    str(manifest_path),
                    "--output",
                    str(bundle),
                ]
            )
            self.assertEqual(0, code)
            self.assertEqual("202601021530", json.loads(bundle.read_text())["run_id"])
            self.assertEqual([], list(bundle.parent.glob(".*.tmp")))

    def test_summary_cli_rejects_existing_output_before_render(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            run_directory = root / "work/review-runs/202601021530"
            run_directory.mkdir(parents=True)
            results_path = run_directory / "results.json"
            results_path.write_text(json.dumps(valid_data("適合")), encoding="utf-8")
            summary = run_directory / "finalize-staging/summary.md"
            summary.parent.mkdir()
            summary.write_text("keep", encoding="utf-8")
            code = MODULE.main(
                [
                    "summary",
                    "--results",
                    str(results_path),
                    "--output",
                    str(summary),
                ]
            )
            self.assertEqual(1, code)
            self.assertEqual("keep", summary.read_text(encoding="utf-8"))

    def test_summary_cli_rejects_output_symlink_ancestor(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary) / "repo"
            run_directory = root / "work/review-runs/202601021530"
            run_directory.mkdir(parents=True)
            results_path = run_directory / "results.json"
            results_path.write_text(json.dumps(valid_data("適合")), encoding="utf-8")
            external = Path(temporary) / "external"
            external.mkdir()
            try:
                (root / "output").symlink_to(external, target_is_directory=True)
            except OSError as exc:
                self.skipTest(f"symlinkを作成できません: {exc}")
            summary = root / "output/reviews/202601021530/summary.md"
            code = MODULE.main(
                [
                    "summary",
                    "--results",
                    str(results_path),
                    "--output",
                    str(summary),
                ]
            )
            self.assertEqual(1, code)
            self.assertEqual([], list(external.iterdir()))

    def test_prepare_rejects_absolute_manifest_entry(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            manifest_path = self._write_prepare_fixture(root)
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            manifest["targets"][0]["original_path"] = str(
                root / "input/targets/design.docx"
            )
            manifest_path.write_text(json.dumps(manifest), encoding="utf-8")
            with self.assertRaisesRegex(MODULE.ReviewDataError, "root相対"):
                MODULE.prepare_bundle(manifest_path)


if __name__ == "__main__":
    unittest.main()
