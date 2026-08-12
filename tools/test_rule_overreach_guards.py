from pathlib import Path
import unittest


ROOT = Path(__file__).resolve().parents[1]
RULES = ROOT / "codex-skill" / "references" / "rules"


def read(path: Path) -> str:
    return path.read_text(encoding="utf-8")


class RuleOverreachGuardTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.skill = read(ROOT / "codex-skill" / "SKILL.md")
        cls.r1 = read(RULES / "r1-cross-document-consistency.md")
        cls.r2 = read(RULES / "r2-required-content.md")
        cls.r3 = read(RULES / "r3-format-and-word-qa.md")
        cls.r4 = read(RULES / "r4-writing-style.md")

    def test_future_use_does_not_require_separate_fixed_end_date(self):
        self.assertIn("将来利用専用の確定保管終期を追記させない", self.r1)
        self.assertIn("通常の保管期間とは別の確定終期がない", self.r4)

    def test_real_future_use_conflicts_remain_reviewable(self):
        self.assertIn("実質的な矛盾がある場合のみ指摘する", self.r1)
        self.assertIn("項目47・説明文書で将来利用の内容・範囲が整合", self.r1)

    def test_inclusive_disposal_policy_is_sufficient(self):
        self.assertIn("包括的な記載で充足", self.r2)
        self.assertIn("廃棄・消去手順は一律に求めない", self.r2)

    def test_missing_disposal_policy_remains_reviewable(self):
        self.assertIn("空欄、未定、または他の保管方針との明確な矛盾", self.r2)

    def test_email_is_not_required_in_participant_information(self):
        self.assertIn("メールアドレスの転記を求めない", self.r1)
        self.assertIn("ある資料にメールアドレスがないこと自体は指摘しない", self.r3)

    def test_reachable_contact_remains_required(self):
        self.assertIn("相談窓口の電話番号と受付時間", self.r2)

    def test_withdrawal_submission_details_are_not_required(self):
        self.assertIn("提出先・提出方法・撤回専用連絡先は一律に求めず", self.r2)
        self.assertIn("同意撤回書が添付され", self.r4)

    def test_required_withdrawal_form_and_limitations_remain_reviewable(self):
        self.assertIn("項目35で文書撤回を選択しているのに同意撤回書がない場合", self.r4)
        self.assertIn("撤回後の措置が困難な場合は、その旨と理由", self.r2)

    def test_skill_surfaces_all_four_guards(self):
        for phrase in (
            "separate fixed end date",
            "media-specific disposal steps",
            "application-system email address",
            "withdrawal-form submission details",
        ):
            self.assertIn(phrase, self.skill)


if __name__ == "__main__":
    unittest.main()
