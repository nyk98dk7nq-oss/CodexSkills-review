import importlib.util
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch


SCRIPT = Path(__file__).parents[1] / "scripts" / "orchestrate.py"
SPEC = importlib.util.spec_from_file_location("orchestrate", SCRIPT)
MODULE = importlib.util.module_from_spec(SPEC)
assert SPEC and SPEC.loader
SPEC.loader.exec_module(MODULE)


class OrchestrateTest(unittest.TestCase):
    def _root(self, temporary):
        root = Path(temporary)
        for relative in (
            "input/checklists",
            "input/references",
            "input/targets",
            "output/reviews",
        ):
            (root / relative).mkdir(parents=True)
        return root

    def test_inventory_rejects_non_xlsx_checklist(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            (root / "input/checklists/check.pdf").write_bytes(b"pdf")
            (root / "input/targets/design.docx").write_bytes(b"docx")
            with self.assertRaisesRegex(MODULE.OrchestrationError, "チェックリストは.xlsx"):
                MODULE.inventory_inputs(root)

    def test_prepare_dispatches_and_preserves_relative_paths(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            self.skipTest("openpyxl is not installed")
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.title = "基本設計"
            worksheet.append(["ID", "チェック項目"])
            worksheet.append([1, "エラー定義があるか"])
            checklist = root / "input/checklists/check.xlsx"
            workbook.save(checklist)
            (root / "input/targets/design.docx").write_bytes(b"docx")
            commands = []

            def fake_runner(command, cwd):
                commands.append(command)
                if "to-markdown" in command:
                    output = Path(command[command.index("to-markdown") + 2])
                    output.parent.mkdir(parents=True, exist_ok=True)
                    if Path(command[1]).name == "image_document.py":
                        output.write_text("OCR座標: x=1, y=2; 信頼度=98\n", encoding="utf-8")
                    else:
                        output.write_text("# converted\n", encoding="utf-8")
                    if Path(command[1]).name == "docx_document.py":
                        images = Path(command[command.index("--images-dir") + 1])
                        images.mkdir(parents=True, exist_ok=True)
                        (images / "embedded.png").write_bytes(b"png")
                elif "prepare" in command and "--output" in command:
                    output = Path(command[command.index("--output") + 1])
                    output.parent.mkdir(parents=True, exist_ok=True)
                    output.write_text("{}\n", encoding="utf-8")

            manifest_path = MODULE.prepare_repository(
                root, "202601021530", runner=fake_runner
            )
            manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(".", manifest["root"])
            self.assertNotIn(str(root), manifest_path.read_text(encoding="utf-8"))
            self.assertEqual("input/checklists/check.xlsx", manifest["checklists"][0]["path"])
            self.assertEqual(
                "work/markdown/202601021530/checklists/check.xlsx.md",
                manifest["checklists"][0]["markdown"],
            )
            self.assertEqual(2, manifest["checklist_items"][0]["row"])
            self.assertIn("チェック項目/B2=エラー定義があるか", manifest["checklist_items"][0]["check_item"])
            scripts = [Path(command[1]).name for command in commands if len(command) > 1]
            self.assertIn("xlsx_document.py", scripts)
            self.assertIn("docx_document.py", scripts)
            self.assertIn("image_document.py", scripts)
            target_markdown = root / "work/markdown/202601021530/targets/design.docx.md"
            self.assertIn("## 抽出画像OCR", target_markdown.read_text(encoding="utf-8"))
            self.assertIn("信頼度=98", target_markdown.read_text(encoding="utf-8"))
            self.assertTrue(
                (root / "work/review-runs/202601021530/review_bundle.json").is_file()
            )

    def test_prepare_rejects_every_existing_run_artifact_before_dispatch(self):
        relative_paths = (
            "work/review-runs/202601021530",
            "work/markdown/202601021530",
            "work/images/202601021530",
            "work/converted-office/202601021530",
            "output/reviews/202601021530",
        )
        for relative in relative_paths:
            with self.subTest(relative=relative), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                conflict = root / relative
                conflict.mkdir(parents=True)

                def unexpected_runner(command, cwd):
                    self.fail("既存run-id検出前にコマンドを実行してはいけません")

                with self.assertRaisesRegex(MODULE.OrchestrationError, "空でも再利用しません"):
                    MODULE.prepare_repository(
                        root, "202601021530", runner=unexpected_runner
                    )

    def test_run_id_rejects_nonexistent_calendar_datetime(self):
        with self.assertRaisesRegex(MODULE.OrchestrationError, "実在するローカル日時"):
            MODULE._validate_run_id("202613321260")

    def test_generated_path_rejects_work_and_output_symlink_ancestors(self):
        for structural in ("work", "work/markdown", "output"):
            with self.subTest(structural=structural), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary) / "repo"
                external = Path(temporary) / "external"
                root.mkdir()
                external.mkdir()
                target = root / structural
                target.parent.mkdir(parents=True, exist_ok=True)
                try:
                    target.symlink_to(external, target_is_directory=True)
                except OSError as exc:
                    self.skipTest(f"symlinkを作成できません: {exc}")
                with self.assertRaisesRegex(MODULE.OrchestrationError, "symlink"):
                    MODULE._reject_existing_run_artifacts(root, "202601021530")
                self.assertEqual([], list(external.iterdir()))

    def test_generated_path_rejects_junction_like_ancestor(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = (Path(temporary) / "repo").resolve()
            junction = root / "work"
            junction.mkdir(parents=True)
            output = junction / "markdown/202601021530/target.md"
            with patch.object(
                type(root),
                "is_junction",
                autospec=True,
                side_effect=lambda candidate: candidate == junction,
            ):
                with self.assertRaisesRegex(
                    MODULE.OrchestrationError, "symlink/junction"
                ):
                    MODULE._validated_repo_path(root, output, "生成先")

    def test_inventory_rejects_input_symlink_outside_repository(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            external = Path(temporary).parent / f"{Path(temporary).name}-external.docx"
            external.write_bytes(b"outside")
            link = root / "input/targets/external.docx"
            try:
                link.symlink_to(external)
            except OSError as exc:
                self.skipTest(f"symlinkを作成できません: {exc}")
            try:
                with self.assertRaisesRegex(MODULE.OrchestrationError, "symlink"):
                    MODULE.inventory_inputs(root)
            finally:
                external.unlink(missing_ok=True)

    def test_results_path_is_exact_relative_path_for_run(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            expected = Path("work/review-runs/202601021530/results.json")
            self.assertEqual(
                root / expected,
                MODULE._expected_results_path(root, "202601021530", expected),
            )
            for invalid in (
                root / expected,
                Path("../results.json"),
                Path("work/review-runs/202601021531/results.json"),
            ):
                with self.subTest(invalid=invalid), self.assertRaisesRegex(
                    MODULE.OrchestrationError, "results"
                ):
                    MODULE._expected_results_path(
                        root,
                        "202601021530",
                        invalid,
                        require_relative=True,
                    )

    def test_manifest_paths_reject_absolute_parent_and_wrong_prefix(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary).resolve()
            base_entry = {
                "path": "input/targets/design.docx",
                "original_path": "input/targets/design.docx",
                "intermediate_path": None,
                "markdown": "work/markdown/202601021530/targets/design.docx.md",
            }
            base = {
                "root": ".",
                "run_id": "202601021530",
                "checklists": [],
                "references": [],
                "targets": [base_entry],
                "checklist_items": [],
            }
            MODULE._validate_manifest_paths(root, "202601021530", base)
            mutations = (
                ("path", str(root / "input/targets/design.docx")),
                ("original_path", "../design.docx"),
                ("markdown", "work/markdown/other/targets/design.md"),
                ("intermediate_path", "../../outside.docx"),
            )
            for key, value in mutations:
                with self.subTest(key=key):
                    manifest = json.loads(json.dumps(base))
                    manifest["targets"][0][key] = value
                    with self.assertRaises(MODULE.OrchestrationError):
                        MODULE._validate_manifest_paths(
                            root, "202601021530", manifest
                        )

    def test_prepare_failure_rolls_back_only_current_run_work(self):
        try:
            from openpyxl import Workbook
        except ImportError:
            self.skipTest("openpyxl is not installed")
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            workbook = Workbook()
            worksheet = workbook.active
            worksheet.append(["ID", "チェック項目"])
            worksheet.append([1, "項目A"])
            workbook.save(root / "input/checklists/check.xlsx")
            (root / "input/targets/design.docx").write_bytes(b"docx")
            unrelated_output = root / "output/reviews/other-run/keep.txt"
            unrelated_output.parent.mkdir(parents=True)
            unrelated_output.write_text("keep", encoding="utf-8")

            def failing_runner(command, cwd):
                if "to-markdown" in command:
                    raise MODULE.OrchestrationError("変換失敗")

            with self.assertRaisesRegex(MODULE.OrchestrationError, "変換失敗"):
                MODULE.prepare_repository(
                    root, "202601021530", runner=failing_runner
                )
            for path in MODULE._run_artifact_paths(root, "202601021530")[:4]:
                self.assertFalse(path.exists(), path)
            self.assertFalse((root / "output/reviews/202601021530").exists())
            self.assertTrue((root / "input/checklists/check.xlsx").is_file())
            self.assertTrue((root / "input/targets/design.docx").is_file())
            self.assertEqual("keep", unrelated_output.read_text(encoding="utf-8"))

            def successful_runner(command, cwd):
                if "to-markdown" in command:
                    output = Path(command[command.index("to-markdown") + 2])
                    output.parent.mkdir(parents=True, exist_ok=True)
                    output.write_text("# converted\n", encoding="utf-8")
                elif "prepare" in command and "--output" in command:
                    output = Path(command[command.index("--output") + 1])
                    output.write_text("{}\n", encoding="utf-8")

            manifest = MODULE.prepare_repository(
                root, "202601021530", runner=successful_runner
            )
            self.assertTrue(manifest.is_file())

    def test_output_names_are_unique_on_case_insensitive_windows(self):
        used = set()
        first = MODULE._unique_output_name("a/Check.xlsx", ".xlsx", used)
        second = MODULE._unique_output_name("b/check.XLSX", ".xlsx", used)
        third = MODULE._unique_output_name("c/CHECK__2.xlsx", ".xlsx", used)
        self.assertEqual("Check.xlsx", first)
        self.assertEqual("check__2.xlsx", second)
        self.assertEqual("CHECK__2__2.xlsx", third)
        self.assertEqual(3, len({first.casefold(), second.casefold(), third.casefold()}))

    def test_format_command_enables_ocr_for_images(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            source = root / "input/targets/screen.png"
            source.write_bytes(b"png")
            output = root / "work/markdown/202601021530/targets/screen.png.md"
            command = MODULE._format_command(
                root, "202601021530", "targets", source, source, output
            )
            self.assertIn("--ocr", command)
            self.assertEqual("jpn+eng+jpn_vert", command[command.index("--lang") + 1])

    def test_unsupported_embedded_image_is_recorded_as_confirmation(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            original = root / "input/targets/design.docx"
            original.write_bytes(b"docx")
            parent_markdown = root / "work/markdown/202601021530/targets/design.docx.md"
            parent_markdown.parent.mkdir(parents=True)
            parent_markdown.write_text("# design\n", encoding="utf-8")
            images = root / "work/images/202601021530/targets/design_docx"
            images.mkdir(parents=True)
            (images / "diagram.emf").write_bytes(b"emf")

            def unexpected_runner(command, cwd):
                self.fail("未対応画像へOCRコマンドを呼んではいけません")

            MODULE._ocr_embedded_images(
                root,
                "202601021530",
                "targets",
                original,
                parent_markdown,
                images,
                unexpected_runner,
            )
            content = parent_markdown.read_text(encoding="utf-8")
            self.assertIn("diagram.emf", content)
            self.assertIn("OCR不可のため要確認。視覚情報を推測しない。", content)

    def test_legacy_frontmatter_points_to_original(self):
        with tempfile.TemporaryDirectory() as temporary:
            markdown = Path(temporary) / "converted.md"
            markdown.write_text(
                '---\nsource_path: "work/converted-office/202601021530/checklists/check.xlsx"\n'
                'source_name: "check.xlsx"\nsource_format: "xlsx"\n---\n# body\n',
                encoding="utf-8",
            )
            MODULE._rewrite_legacy_frontmatter(
                markdown,
                "input/checklists/check.xls",
                "work/converted-office/202601021530/checklists/check.xlsx",
                ".xls",
                ".xlsx",
            )
            result = markdown.read_text(encoding="utf-8")
            self.assertIn('source_path: "input/checklists/check.xls"', result)
            self.assertIn('source_format: "xls"', result)
            self.assertIn(
                'intermediate_path: "work/converted-office/202601021530/checklists/check.xlsx"',
                result,
            )

    def test_results_must_match_manifest(self):
        manifest = {
            "checklists": [{"path": "input/checklists/check.xlsx"}],
            "references": [],
            "targets": [{"path": "input/targets/a.docx"}],
            "checklist_items": [
                {
                    "checklist_file": "input/checklists/check.xlsx",
                    "sheet": "S",
                    "row": 2,
                    "check_item": "項目A",
                }
            ],
        }
        results = json.loads(json.dumps(manifest))
        MODULE._verify_results_match_manifest(manifest, results)
        results["targets"] = [{"path": "input/targets/other.docx"}]
        with self.assertRaisesRegex(MODULE.OrchestrationError, "manifestと一致"):
            MODULE._verify_results_match_manifest(manifest, results)

    def test_finalize_publishes_staging_only_after_all_outputs_exist(self):
        with tempfile.TemporaryDirectory() as temporary:
            root = self._root(temporary)
            checklist = root / "input/checklists/check.xlsx"
            checklist.write_bytes(b"xlsx")
            markdown = "work/markdown/202601021530/checklists/check.xlsx.md"
            target_markdown = "work/markdown/202601021530/targets/design.docx.md"
            manifest = {
                "run_id": "202601021530",
                "root": ".",
                "checklists": [
                    {
                        "path": "input/checklists/check.xlsx",
                        "original_path": "input/checklists/check.xlsx",
                        "markdown": markdown,
                        "header_row": 1,
                    }
                ],
                "references": [],
                "targets": [
                    {
                        "path": "input/targets/design.docx",
                        "original_path": "input/targets/design.docx",
                        "markdown": target_markdown,
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
            run_directory = root / "work/review-runs/202601021530"
            run_directory.mkdir(parents=True)
            (run_directory / "manifest.json").write_text(
                json.dumps(manifest), encoding="utf-8"
            )
            results = {
                **manifest,
                "generated_at": "2026-01-02T15:30:00+09:00",
                "items": [
                    {
                        **manifest["checklist_items"][0],
                        "target_file": "input/targets/design.docx",
                        "result": "適合",
                        "comment": "target.md:L1に定義がある。",
                        "evidence": ["target.md:L1"],
                        "improvement": "",
                    }
                ],
            }
            results_path = run_directory / "results.json"
            results_path.write_text(json.dumps(results), encoding="utf-8")

            def failing_runner(command, cwd):
                if "summary" in command:
                    output = Path(command[command.index("--output") + 1])
                    output.write_text("# summary\n", encoding="utf-8")
                elif "write-review" in command:
                    raise MODULE.OrchestrationError("結果書込失敗")

            with self.assertRaisesRegex(MODULE.OrchestrationError, "結果書込失敗"):
                MODULE.finalize_repository(
                    root,
                    "202601021530",
                    Path("work/review-runs/202601021530/results.json"),
                    runner=failing_runner,
                )
            self.assertFalse((run_directory / "finalize-staging").exists())
            self.assertFalse((root / "output/reviews/202601021530").exists())
            self.assertTrue(checklist.is_file())
            self.assertTrue(results_path.is_file())

            def fake_runner(command, cwd):
                if "summary" in command:
                    output = Path(command[command.index("--output") + 1])
                    output.write_text("# summary\n", encoding="utf-8")
                elif "write-review" in command:
                    output = Path(command[command.index("write-review") + 2])
                    output.write_bytes(b"result")

            output = MODULE.finalize_repository(
                root,
                "202601021530",
                Path("work/review-runs/202601021530/results.json"),
                runner=fake_runner,
            )
            self.assertTrue((output / "summary.md").is_file())
            self.assertTrue((output / "check.xlsx").is_file())
            self.assertFalse((run_directory / "finalize-staging").exists())


if __name__ == "__main__":
    unittest.main()
