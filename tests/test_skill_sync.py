import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
from sync_precheck_skill import synchronize


class SkillSyncTests(unittest.TestCase):
    def test_package_mirrors_all_maintained_sources(self):
        result = synchronize(check=True)
        self.assertEqual(result["differences"], 0, result["changes"])


if __name__ == "__main__":
    unittest.main()
