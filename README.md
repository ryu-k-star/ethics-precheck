# ethics-precheck

日本の臨床研究・倫理審査申請資料を、Codexで根拠に沿ってプレチェックするための公開 Skill です。PDF/DOCX の申請資料一式を対象に、51項目の確認、文書間の整合性監査、指摘根拠の記録、施設様式を保った返却用 Word の作成を支援します。

この Skill は倫理審査委員会、事務局、法務・コンプライアンス担当者の判断を代替しません。出力は人による確認を前提としたドラフトです。

## 主な機能

- 申請資料、研究計画書、説明・同意文書、オプトアウト文書、質問票などを横断して確認
- 申請51項目と関連文書の不一致、記載不足、書式残骸をルールID単位で監査
- 初回申請と再申請の両方に対応し、過去指摘に引きずられないゼロベース確認を実施
- 根拠資料とページ・項目を付けた Markdown の監査記録を作成
- 同梱テンプレートの表・枠を保った返却用 DOCX を生成し、転記漏れを検証
- 不明点、施設判断、根拠不足を断定せず「人間確認事項」として分離

## インストール

Skill 本体はリポジトリ直下ではなく [`codex-skill/`](codex-skill/) にあります。

### Codex に依頼する方法（推奨）

Codex で次のように依頼してください。

```text
$skill-installer を使って、次の GitHub URL から Skill をインストールしてください。
https://github.com/ryu-k-star/ethics-precheck/tree/master/codex-skill
```

インストール後は、次のターンまたは新しいタスクから `$ethics-precheck` を利用できます。

### 手動で配置する方法

リポジトリを取得し、`codex-skill` ディレクトリを Codex の Skills ディレクトリへ `ethics-precheck` という名前でコピーします。

PowerShell:

```powershell
git clone --depth 1 https://github.com/ryu-k-star/ethics-precheck.git
New-Item -ItemType Directory -Force "$env:USERPROFILE\.codex\skills" | Out-Null
Copy-Item -Recurse ".\ethics-precheck\codex-skill" "$env:USERPROFILE\.codex\skills\ethics-precheck"
```

macOS / Linux:

```bash
git clone --depth 1 https://github.com/ryu-k-star/ethics-precheck.git
mkdir -p "${CODEX_HOME:-$HOME/.codex}/skills"
cp -R ./ethics-precheck/codex-skill "${CODEX_HOME:-$HOME/.codex}/skills/ethics-precheck"
```

同名ディレクトリがすでにある場合は、内容を確認してから更新してください。

## 要件

- Skills を利用できる Codex 環境
- Python 3.9 以上
- Python パッケージ: `pdfplumber`、`python-docx`
- PDF/DOCX の読み書きが許可されたローカル作業環境

補助スクリプトの依存関係は、インストール後に次で導入できます。

```powershell
python -m pip install -r "$env:USERPROFILE\.codex\skills\ethics-precheck\requirements.txt"
```

macOS / Linux では `$CODEX_HOME/skills/ethics-precheck/requirements.txt`（`CODEX_HOME` 未設定時は `~/.codex/skills/ethics-precheck/requirements.txt`）を指定してください。

DOCX の全ページを画像または PDF にレンダリングできる環境があると、構造確認に加えて目視確認まで行えます。レンダリング手段がない場合、Skill はその制約を明記し、視覚的QAを完了扱いにしません。

## 使用例

初回申請:

```text
$ethics-precheck を使って C:\research\2026-001 の倫理申請資料をプレチェックしてください。
根拠箇所を付けた Markdown 一式と、施設テンプレートを保持した返却用 Word のドラフトを作成してください。
```

再申請:

```text
$ethics-precheck を使って、この再申請を51項目すべてゼロベースで確認してください。
前回指摘への対応状況は独立して照合し、現在資料に根拠がない指摘は採用しないでください。
```

視覚的QAを含む確認:

```text
$ethics-precheck で返却用 Word を作成し、可能なら全ページをレンダリングして表崩れ、切れ、転記位置も確認してください。
```

## 入力と出力

想定する入力は、研究ごとのローカルフォルダに置かれた申請書、研究計画書、説明・同意文書、オプトアウト文書、質問票、データ・試料取扱文書、前回指摘書などです。

標準出力は、研究フォルダ内の新しい `99_プレチェック出力/` に保存します。

```text
01_資料一覧.md
02_研究種別_仮分類.md
03_サブAI別レビュー結果.md
04_指摘事項書ドラフト.md
05_人間確認事項.md
06_ファクトチェック結果.md
★倫理申請_修正対応依頼_<研究課題名>_<作成日4桁>.docx
```

既存の原資料、過去出力、同梱テンプレートは上書きしません。

## データとプライバシー

このリポジトリには研究資料や実案件の出力を含めないでください。研究資料には個人情報・要配慮個人情報・未公開研究情報が含まれる可能性があります。

- 施設の情報管理規程と、利用する Codex 環境のデータ取扱条件を確認してください。
- 研究資料、抽出テキスト、生成したレビュー結果を GitHub へコミットしないでください。
- 公開リポジトリへ Issue やログを投稿するときは、研究名、氏名、連絡先、症例情報などを除去してください。
- 法令、倫理指針、施設内規程の最新性が判断に影響する場合は、公式の最新版を確認してください。

`.gitignore` は典型的な研究フォルダや一時出力を除外しますが、すべての命名を自動で検出できるわけではありません。コミット前の確認は利用者の責任で行ってください。

## リポジトリ構成

```text
codex-skill/                 配布する Codex Skill
  SKILL.md                   発動条件と中核ワークフロー
  agents/openai.yaml         Codex UI メタデータ
  assets/                    返却用 Word テンプレート
  references/                詳細ワークフロー、ルール、品質基準
  scripts/                   抽出・DOCX生成スクリプト
rules/                       リポジトリ運用版の日本語ルールブック
docs/                        品質管理資料
tools/                       開発・検証用ツールとテスト
AGENTS.md                    リポジトリ内での詳細運用ルール
```

## 開発者向け検証

```powershell
python -m unittest discover -s tools -p "test_*.py"
python -m py_compile codex-skill/scripts/extract_research_docs.py codex-skill/scripts/build_return_docx_from_template.py
```

Skill の frontmatter と構成は、Codex に同梱される `quick_validate.py` でも検証できます。

## ライセンスと免責

このリポジトリは [MIT License](LICENSE) で公開します。

本ソフトウェアおよび文書は現状有姿で提供され、特定目的への適合性、正確性、法令・指針への適合、審査結果を保証しません。これは医療・法律・倫理審査上の助言、承認、不承認、適合証明ではありません。研究者への返却や申請への使用前に、権限を有する人が原資料、根拠、表現、最新の法令・指針・施設内規程を確認してください。
