# Badminton Scheduler

バドミントンの組み合わせ（対戦表）を、同レベル・同グループ回避などの条件を守りつつ自動生成し、対戦表（Excel）と閲覧用HTMLを出力するツールです。

## できること（運用でよく使う指定）

- コート数を指定して生成（`--courts`）
- ラウンド数（=全体の試合枠）を指定して生成（`--num-rounds`）
- 各ペアの試合数を「自動」または「固定」で指定（`--matches-per-team`）
  - `0` = 自動（全員同じ試合数を最優先し、容量内で成立する最大値を選びます）
  - `6` = 固定6試合、など
- 「この対戦は入れたい」を事前に指定（入力Excelの `優先1〜3` / `希望` / `対戦相手` などの列）

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

初見の人は、まず [工程表.md](工程表.md)（= 運用の流れ）だけ読むと迷いにくいです。

ざっくり流れ（よくある2パターン）:

- チームリスト（xlsx）を作る → 自動生成 → （必要ならExcelで少し修正）→ 修正版を入力して配布物を作り直す
- 集計表（xlsm）の `対戦一覧_短縮` が正 → そこから配布物を一発で再生成

## Streamlit（アップロード/デプロイ）

このリポジトリは Streamlit Community Cloud を想定した構成です。

- エントリポイント: [streamlit_app.py](streamlit_app.py)（中で [app.py](app.py) を読み込みます）
- 依存関係: [requirements.txt](requirements.txt)
- 合言葉（必須）: Secrets に `APP_PASSCODE` を設定
  - 例: [.streamlit/secrets.toml.example](.streamlit/secrets.toml.example)

注意（個人情報/公開範囲）:

- Streamlit Community Cloud に置いたアプリは「Web公開」です。URLを知っている人がアクセスできます。
- `APP_PASSCODE` は簡易ロックで、暗号化ではありません（強い秘匿が必要な情報の扱いには向きません）。
- 実名などが入るExcelを扱う場合は、公開範囲や運用（共有先）を必ず決めてください。

手順（概要）:

1) このフォルダを GitHub に push
2) Streamlit Cloud で「New app」→ リポジトリ選択 → Main file path に `streamlit_app.py`
3) Settings → Secrets に `APP_PASSCODE = "...."` を登録

初回デプロイ後の使い方（ブラウザだけで完結する想定）:

1) チームリスト（xlsx）をアップロードして生成
2) 出てきた `schedule.xlsx` をダウンロード
3) 現場で少し入れ替えたら、その修正版 `schedule.xlsx` をアップロードして配布HTML/壁貼り/得点記入表を作り直す

### 1) まずヘルプ確認

```powershell
bsched --help
```

（リポジトリ直下で実行する場合は `python -m badminton_program.scheduler --help` でもOKです）

### 2) 例: 15面・23ラウンド・6試合固定で生成

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6
```

現場で手修正しやすくする（対戦表のセルに「ペア名 + 改行 + 選手名」を入れる）:

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6 --excel-include-members
```

別案（名前変更が多い現場向け）: **1ラウンド=2行**（上=ペア名、下=選手名）にして、下段は `ペア一覧` から VLOOKUP で自動表示:

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6 --excel-members-below --excel-members-vlookup
```

- この形だと、氏名変更は基本 `ペア一覧` を直すだけでOK（対戦表はペア名の入れ替えに集中できます）
- VLOOKUP は Excel で開いて計算された結果が見える想定です（必要なら一度開いて保存してください）

### 2.5) 先に「入れたい対戦」を指定してから生成（おすすめ）

対戦表をExcel上で大量に入れ替えるのはミスりやすいので、
**「入れたい対戦（固定したい相手）」は先に `チームリスト.xlsx` 側へ書いて、生成をやり直す**のが安定します。

入力Excel（`チームリスト.xlsx`）の各行にある `優先1` / `優先2` / `優先3`（または列名に `優先` / `希望` / `対戦` / `相手` を含む列）へ、
「当てたい相手のペア名」を書くと、その組み合わせを優先して入れようとします。

ポイント:

- 片方だけに書いてもOKです（A行に「優先1=B」と書けば A vs B を希望として扱います）。
- ただし、ルール上できないものは入りません（例: 同レベルのみ／同グループは回避）。
- 希望を入れすぎると成立しなくなることがあります（各ペアの試合数上限を超える希望など）。

同じ希望条件のまま「別の並び」を見たい時は、`--diversity-attempts` を増やすと（同条件で）複数回試行します。

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6 --diversity-attempts 10
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

## 現実運用: Excelを手修正 → 再出力

実運用では「先に組合せを作る→微修正が入る→当日朝に氏名が変わる」ことが多いので、
このツールは **編集済みExcelを入力にして HTML/Excel を再生成**できます。

使い分けの目安:

- 大きく組合せ方針を変える（入れたい対戦が多い）: `チームリスト.xlsx` に希望対戦を書いて `bsched` で生成し直す
- 当日現場の微修正（数試合だけ差し替え）: 生成済みExcelの `対戦表` を手で触り、`html-from-xlsx` / `xlsx-from-xlsx` で吐き直す

### 名前変更が多い場合（おすすめ運用: ペア一覧を同期）

「氏名（メンバー名）のマスターが別Excel（集計表/xlsm）にある」場合は、
**対戦表を直すExcelとは別に、氏名を1か所で管理**したくなります。

このツール側では、スケジュールExcelの `ペア一覧` を「最新チーム一覧」から同期できます。

流れ（例）:

1) 集計表（xlsm）側で氏名を更新
2) 集計表 → `チームリスト.xlsx` を作る（例: [create_team_list.py](create_team_list.py)）

```powershell
python create_team_list.py --input-file "集計表.xlsm" --output-file "チームリスト.xlsx" --sheet "リスト"
```

3) （手修正済みの）スケジュールExcelへ `ペア一覧` を同期

```powershell
python -m badminton_program.scheduler sync-pairs-from-team-list --schedule-file "local_out\schedule_edit.xlsx" --team-list-file "チームリスト.xlsx"
```

4) 最終配布物を吐き直す

```powershell
python -m badminton_program.scheduler xlsx-from-xlsx --input-file "local_out\schedule_edit_pairsynced.xlsx" --output-file "local_out\schedule_final.xlsx" --excel-members-below --excel-members-vlookup
python -m badminton_program.scheduler html-from-xlsx --input-file "local_out\schedule_final.xlsx" --include-members --wall-html --wall-courts-per-page 3
```

この運用だと、氏名は基本「集計表→チームリスト→同期」で一本化でき、
対戦表（組合せ）を壊さずに最新の氏名だけ反映できます。

### 生成→手修正→集計表へ流し込み→マクロ短縮→最終出力（あなたの想定フロー）

「このツールで組合せを作ってExcelに吐く → 現場で少し入れ替える → その結果を集計表(xlsm)へ入れる → 集計表マクロで `対戦一覧_短縮` を作る → それをPythonが読んで最終HTML/Excelを出す」
という流れは **この順番が正** です。

1) まずスケジュールを生成（例）

```powershell
bsched --input-file "チームリスト.xlsx" --output-file "schedule.xlsx" --courts 15 --num-rounds 23 --matches-per-team 6
```

2) `schedule.xlsx` をExcelで手修正（必要なら）

3) 修正版 `schedule.xlsx` を、集計表(xlsm)の「3行=1試合」入力グリッドへ流し込み

```powershell
python -m badminton_program.scheduler fill-summary-grid-from-xlsx \
  --schedule-file "schedule.xlsx" \
  --summary-file "集計表.xlsm"
```

4) できた `*_filled_from_xlsx.xlsm` をExcelで開いて、マクロで `対戦一覧_短縮` を再生成

5) その集計表を入力にして、最終配布物を出力

```powershell
python -m badminton_program.scheduler export-from-summary --input-file "集計表_filled_from_xlsx.xlsm" --wall-html --wall-courts-per-page 3
python -m badminton_program.scheduler score-sheets-from-summary --input-file "集計表_filled_from_xlsx.xlsm"
```

※ポイント: `fill-summary-grid` は「`対戦一覧_短縮` → 集計表グリッド」用です（短縮一覧を正にする運用向け）。
この章のフロー（スケジュールExcelを正にする）では、`fill-summary-grid-from-xlsx` を使います。

### 迷わない運用（おすすめ: マスターは1ファイルだけ）

現場で混乱しない一番簡単なルールはこれです:

- **マスターは常に集計表(xlsm) 1つ**（例: `*_filled_from_xlsx.xlsm`）
- 手修正は Excel だけでOK（最終的にマスターへ反映されていれば良い）
- 配布物（HTML/Excel/得点票）は **必ずマスターから出す**

Excelだけの人に渡す時の実務フロー（最小ステップ）:

1) （あなたが）スケジュールを作って `schedule.xlsx` を渡す / または自分で手修正
2) （あなたが）`scripts/01_make_master_from_schedule.ps1` を実行してマスター(xlsm)を作る
3) （Excel作業）マスターを開き、マクロで `対戦一覧_短縮` を更新
4) （あなたが）`scripts/02_export_from_master.ps1` を実行して配布物を出す

スクリプト例（PowerShell）:

```powershell
./scripts/01_make_master_from_schedule.ps1 -ScheduleFile "schedule.xlsx" -SummaryFile "集計表.xlsm"
./scripts/02_export_from_master.ps1 -MasterSummaryFile "集計表_filled_from_xlsx.xlsm" -WallHtml
```

### オンラインでやりたい場合（重要: マクロはクラウドで動きません）

Streamlit Cloud / GitHub Actions などの「オンライン実行」では、Excelマクロ（xlsmのボタン）を実行できません。
そのため、オンラインで完結させるなら **`schedule.xlsx` を正として、Pythonで配布物を作る**のが一番簡単です。

- `app.py`（Streamlit）では
  - チーム一覧Excelアップロード → 生成（Excel + HTML）
  - 編集後 `schedule.xlsx` アップロード → HTML再生成（+ 壁貼りHTML / 得点記入表HTML）
 までをブラウザだけでできます。

集計表(xlsm)を使う運用（得点入力・集計など）をしたい場合は、
"オンラインで配布物" と "ローカルExcelで得点" を分けるのが現実的です。

### さらに簡単: 集計表（xlsm）の「対戦一覧_短縮」から直接出力

すでに集計表（xlsm）側に `対戦一覧_短縮` シートがあり、
そこに「試合/コート/時間/ペア名/選手名/相手ペア名/相手選手名」が揃っているなら、
**そのxlsmを入力にして、最終Excel+HTMLを直接作る**のが一番混乱が少ないです。

```powershell
python -m badminton_program.scheduler export-from-summary --input-file "集計表.xlsm" --sheet-name "対戦一覧_短縮" --wall-html --wall-courts-per-page 3
```

（必要に応じて）Excelの対戦表を見やすい2行レイアウトで出す:

```powershell
python -m badminton_program.scheduler export-from-summary --input-file "集計表.xlsm" --sheet-name "対戦一覧_短縮" --excel-members-below --excel-members-vlookup
```

### A) 編集したExcelからHTMLを再生成

（組合せを手で入れ替えた／氏名を直した後に、配布用HTMLを作り直す）

ポイント:

- `--excel-include-members` のExcelは、`対戦表` のセルに氏名も入るので、そのシートだけ見ながら（コピー&ペーストで）組合せ調整しやすいです。
  - `html-from-xlsx` は `対戦表` セル内の氏名（改行2行目以降）も拾います（= `ペア一覧` を更新しなくても反映される場合があります）
- `--excel-members-below --excel-members-vlookup` のExcelは、氏名変更は基本 `ペア一覧` を更新し、対戦表はペア名の入れ替えに集中できます。

```powershell
python -m badminton_program.scheduler html-from-xlsx --input-file "local_out\schedule_edit.xlsx" --include-members
```

※ `Permission denied` が出る場合は、対象Excelを開いたままのことが多いです（Excelを閉じて再実行）。

壁貼りHTMLも同時に作る場合:

```powershell
python -m badminton_program.scheduler html-from-xlsx --input-file "local_out\schedule_edit.xlsx" --include-members --wall-html --wall-courts-per-page 3
```

### B) 編集したExcelからExcelを作り直す（短縮表/個人表の再生成）

編集は `対戦表` や `ペア一覧` を触ることが多いですが、短縮表・個人表は自動更新されません。
そのため「編集済みExcel → 派生シートを再計算したExcel」を作りたい時はこれを使います。

```powershell
python -m badminton_program.scheduler xlsx-from-xlsx --input-file "local_out\schedule_edit.xlsx" --output-file "local_out\schedule_final.xlsx"
```

※ `Permission denied` が出る場合は、対象Excelを閉じてから実行してください。

その後、最終版HTMLを作る:

```powershell
python -m badminton_program.scheduler html-from-xlsx --input-file "local_out\schedule_final.xlsx" --include-members --wall-html --wall-courts-per-page 3
```

## 入力ファイル

- `チームリスト.xlsx` を入力にします（列名などはサンプルを参照してください）。

現在の推奨テンプレは **4列だけ** です:

- `ペア名`, `氏名`, `優先対戦`, `優先対戦相手`

互換: 旧テンプレの列名（例: `ペア名 ↓値ばりで記入`, `氏名 ↓値ばりで記入`）でも読み込めます。

`レベル` / `グループ` 列は不要です（ペア名から自動推定します。レベルは原則として末尾の数字の直前にある `A/B/C`（例: `上海A1`）を使い、グループは末尾数字を除いた部分を使います）。

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
