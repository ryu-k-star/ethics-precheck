import json
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class ScanTests(unittest.TestCase):
    def test_cache_and_output_exclusion(self):
        with tempfile.TemporaryDirectory() as temp:
            base = Path(temp)
            study = base / "study"
            study.mkdir()
            (study / "source.txt").write_text("研究期間 2026年1月1日から2027年3月31日 統計解析 3か月 200例", encoding="utf-8")
            excluded = study / "99_プレチェック出力_old"
            excluded.mkdir()
            (excluded / "old.pdf").write_bytes(b"not-a-pdf")
            out = base / "out"
            cache = base / "cache"
            cmd = [sys.executable, str(ROOT / "tools/precheck_scan.py"), str(study), "--legacy-layout", "--out", str(out), "--cache", str(cache), "--rules", str(ROOT / "rules")]
            first = subprocess.run(cmd, check=False, capture_output=True, text=True)
            self.assertEqual(first.returncode, 0, first.stderr)
            second = subprocess.run(cmd, check=False, capture_output=True, text=True)
            self.assertEqual(second.returncode, 0, second.stderr)
            self.assertTrue(json.loads(second.stdout)["reused_existing_output"])
            # A new review output reuses extraction; the old review is never reset.
            cmd[cmd.index("--out") + 1] = str(base / "out2")
            third = subprocess.run(cmd, check=False, capture_output=True, text=True)
            self.assertEqual(third.returncode, 0, third.stderr)
            out = base / "out2"
            manifest = json.loads((out / "00_source_manifest.json").read_text(encoding="utf-8"))
            self.assertEqual([f["path"] for f in manifest["files"]], ["source.txt"])
            self.assertTrue(manifest["files"][0]["cache_hit"])
            self.assertTrue(any("old.pdf" in warning for warning in manifest["warnings"]))
            timeline = json.loads((out / "00_timeline_facts.json").read_text(encoding="utf-8"))
            self.assertEqual(timeline["evidence"][0]["targets"], [200])
            ledger = json.loads((out / "00_rule_ledger.json").read_text(encoding="utf-8"))
            self.assertEqual(ledger["schema_version"], 2)
            self.assertIn("researcher_profile", ledger["allowed_return_routes"])
            self.assertEqual(ledger["dispositions"], [])


if __name__ == "__main__":
    unittest.main()
