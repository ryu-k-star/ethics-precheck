from __future__ import annotations

import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def normalized(path: Path) -> str:
    return path.read_text(encoding="utf-8").replace("\r\n", "\n").rstrip()


class SkillPackageMirrorTests(unittest.TestCase):
    def test_release_mirrors_match_repository_sources(self) -> None:
        mappings = {
            "rules/R1_項目間整合ルール.md": "codex-skill/references/rules/r1-cross-document-consistency.md",
            "rules/R2_必須記載チェックリスト.md": "codex-skill/references/rules/r2-required-content.md",
            "rules/R3_書式残骸・表記ルール.md": "codex-skill/references/rules/r3-format-and-word-qa.md",
            "rules/R4_文体規範と文例集.md": "codex-skill/references/rules/r4-writing-style.md",
            "rules/R5_事務局定型文.md": "codex-skill/references/rules/r5-office-standard-text.md",
            "docs/review_quality_playbook.md": "codex-skill/references/review_quality_playbook.md",
            "tools/build_return_docx_from_template.py": "codex-skill/scripts/build_return_docx_from_template.py",
            "tools/precheck_scan.py": "codex-skill/scripts/precheck_scan.py",
            "tools/precheck_verify.py": "codex-skill/scripts/precheck_verify.py",
            "tools/precheck_usage.py": "codex-skill/scripts/precheck_usage.py",
        }
        mismatches = [
            f"{source} != {mirror}"
            for source, mirror in mappings.items()
            if normalized(ROOT / source) != normalized(ROOT / mirror)
        ]
        self.assertEqual(mismatches, [])


if __name__ == "__main__":
    unittest.main()
