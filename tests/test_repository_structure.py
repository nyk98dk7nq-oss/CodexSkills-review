from __future__ import annotations

import subprocess
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
SKILL_NAMES = {
    "xlsx-document",
    "docx-document",
    "pptx-document",
    "pdf-document",
    "image-document",
    "convert-legacy-office",
    "review-markdown-documents",
    "review-documents-orchestrator",
}


class RepositoryStructureTests(unittest.TestCase):
    def test_required_root_files_exist(self) -> None:
        for relative in (
            "README.md",
            "AGENTS.md",
            ".gitignore",
            "requirements.txt",
            "review.py",
            "test_all.py",
        ):
            self.assertTrue((ROOT / relative).is_file(), relative)

    def test_user_facing_directories_exist(self) -> None:
        for relative in (
            "input/checklists/README.md",
            "input/references/README.md",
            "input/targets/README.md",
            "work/README.md",
            "output/reviews/README.md",
            "output/edited/README.md",
        ):
            self.assertTrue((ROOT / relative).is_file(), relative)

    def test_expected_skill_directories_exist(self) -> None:
        skill_root = ROOT / ".agents" / "skills"
        actual = {path.name for path in skill_root.iterdir() if path.is_dir()}
        self.assertEqual(SKILL_NAMES, actual)

    def test_skill_metadata_and_scripts_exist(self) -> None:
        for name in sorted(SKILL_NAMES):
            with self.subTest(skill=name):
                skill_dir = ROOT / ".agents" / "skills" / name
                skill_md = skill_dir / "SKILL.md"
                metadata = skill_dir / "agents" / "openai.yaml"
                scripts = list((skill_dir / "scripts").glob("*.py"))
                self.assertTrue(skill_md.is_file())
                self.assertTrue(metadata.is_file())
                self.assertTrue(scripts)
                text = skill_md.read_text(encoding="utf-8")
                self.assertIn(f"name: {name}", text)
                self.assertNotIn("TODO", text)
                self.assertIn(f"${name}", metadata.read_text(encoding="utf-8"))

    def test_review_launcher_help(self) -> None:
        completed = subprocess.run(
            [sys.executable, str(ROOT / "review.py"), "--help"],
            capture_output=True,
            text=True,
            check=False,
        )
        self.assertEqual(0, completed.returncode, completed.stderr)
        self.assertIn("preflight", completed.stdout)


if __name__ == "__main__":
    unittest.main()
