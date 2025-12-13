# Badminton Scheduler

バドミントンの組み合わせ（対戦表）を、同レベル・同グループ回避などの条件を守りつつ自動生成し、対戦表（Excel）と閲覧用HTMLを出力するツールです。

## できること（運用でよく使う指定）

- コート数を指定して生成（`--courts`）
- ラウンド数（=全体の試合枠）を指定して生成（`--num-rounds`）
- 各ペアの試合数を「自動」または「固定」で指定（`--matches-per-team`）
  - `0` = 自動（全員同じ試合数を最優先し、容量内で成立する最大値を選びます）
  - `6` = 固定6試合、など

※容量の考え方: `容量 = num_rounds * courts`、必要試合数は `必要 = (ペア数 * 目標試合数) / 2` です。

## セットアップ

Python 3.10+ が必要です。

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -U pip
pip install -e .
```

## 使い方

### 1) まずヘルプ確認

```powershell
bsched --help
```

（リポジトリ直下で実行する場合は `python -m badminton_program.scheduler --help` でもOKです）

### 2) 例: 15面・23ラウンド・6試合固定で生成

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6
```

### 3) 例: 各ペア試合数は自動（0）で生成

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 0
```

### 4) HTMLを簡易ロック付きで生成（静的HTML）

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6 --html-passcode "1234"
```

注意: これは「簡易ロック」です（完全な暗号化ではありません）。

## 入力ファイル

- `チームリスト.xlsx` を入力にします（列名などはサンプルを参照してください）。

## 公開（GitHub）について

- `pyproject.toml` の `license` は現在 `Proprietary` になっています。公開するなら、意図に合わせてライセンス表記の見直しが必要です。
- 個人情報が入る可能性があるため、入力Excelや出力HTML/Excelをそのままリポジトリに含めない運用を推奨します（`.gitignore`推奨）。

補足: GitHubを `Private` にすると、招待したメンバー以外は見られません（=「みんなに公開」はできません）。

## GitHub上で入力して実行（Actions）

GitHubの画面から「コート数・ラウンド数・各ペア試合数」などを入力して実行し、生成されたExcel/HTMLをArtifactsからダウンロードできます。

1. リポジトリに `チームリスト.xlsx` を置く（または入力パスを変更）
2. GitHub → `Actions` → `Generate Schedule` → `Run workflow`
3. 入力欄に `courts`, `num_rounds`, `matches_per_team` などを入れて実行
4. 実行結果ページの `Artifacts` から `schedule-output` をダウンロード

## Streamlitで回す（ブラウザUI）

このリポジトリには簡易Streamlitアプリ [app.py](app.py) を含めています。

合言葉（簡易アクセス制限）:
- デフォルトは `1234` です。
- Streamlit Cloud の `Secrets` に `APP_PASSCODE` を設定すると変更できます。

ローカル起動:

```powershell
pip install -e .
streamlit run app.py
```

デプロイ:
- Streamlit Community Cloudは基本的に「GitHubのPublicリポジトリ」が前提です。
- Privateのまま使うなら、Streamlit Cloudの有料プランか、自前サーバー/自PCでの起動が現実的です。

### Streamlit Community Cloud で公開する手順（概要）

1. GitHubのリポジトリを `Public` にする（Community Cloudの前提）
2. Streamlit Community Cloud で `New app` → 対象リポジトリ/ブランチを選択
3. `Main file path` を `app.py` にする
4. デプロイ

このリポジトリには `requirements.txt` を入れてあり、`-e .` でパッケージをインストールします（src構成でも import できるようにするため）。

HTMLはサーバー上で公開せず、生成後にダウンロードさせる運用を推奨します（簡易ロックは暗号化ではありません）。
