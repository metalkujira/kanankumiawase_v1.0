from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from badminton_program import scheduler


st.set_page_config(page_title="Badminton Scheduler", layout="wide")

st.title("Badminton Scheduler")
st.caption("Excelのチームリストをアップロードして、コート数/ラウンド数/各ペア試合数を指定して生成します。")
st.warning(
    "公開運用の注意: アップロードされたExcelはサーバー側で処理されます。アプリ側では一時ディレクトリで処理して終了後に削除しますが、基盤（Streamlit Cloud等）のログ/監視/バックアップ等まで含めて完全な削除を保証はできません。個人情報は匿名化/最小化してからアップロードしてください。",
)


def _expected_app_passcode() -> str:
    # Default is '1234' as a minimal gate; override via Streamlit Secrets (APP_PASSCODE).
    try:
        return str(st.secrets.get("APP_PASSCODE", "1234"))
    except Exception:
        return "1234"


if "authed" not in st.session_state:
    st.session_state.authed = False

expected_passcode = _expected_app_passcode()

with st.sidebar:
    st.header("アクセス")
    entered = st.text_input("合言葉", value="", type="password")
    if st.button("解除", use_container_width=True):
        st.session_state.authed = (entered == expected_passcode)
        if not st.session_state.authed:
            st.error("合言葉が違います")

if not st.session_state.authed:
    st.info("合言葉を入力して解除してください")
    st.stop()

with st.sidebar:
    st.header("設定")

    st.download_button(
        "テンプレExcel(ヘッダーのみ)をダウンロード",
        data=scheduler.build_team_list_template_bytes(),
        file_name="チームリスト_テンプレ.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )
    st.caption("このテンプレに入力してアップロードしてください（先頭/左端のシートを読みます）。")

    st.markdown("### サンプル（ダミーデータ）")
    st.caption("適当なチーム名・氏名が入ったサンプルExcelをダウンロードできます（個人情報なし）")
    sample_bytes = scheduler.build_team_list_sample_bytes()
    st.download_button(
        label="チーム一覧サンプル（ダミーデータ）をダウンロード",
        data=sample_bytes,
        file_name="チームリスト_サンプル.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    uploaded = st.file_uploader("チーム一覧Excel (.xlsx)", type=["xlsx"]) 

    courts = st.number_input("コート数", min_value=1, max_value=60, value=15, step=1)
    num_rounds = st.number_input("ラウンド数", min_value=1, max_value=200, value=23, step=1)
    start_time = st.text_input("開始時刻 (HH:MM)", value="12:50")
    round_minutes = st.number_input("1ラウンドの時間（分）", min_value=1, max_value=180, value=13, step=1)
    matches_per_team = st.number_input("各ペアの試合数 (0=自動)", min_value=0, max_value=30, value=0, step=1)

    max_consecutive = st.selectbox(
        "最大連戦数（基本2。無理なら自動で3）",
        options=[2, 3],
        index=0,
    )

    diversity_attempts = st.number_input("分散最大化の試行回数", min_value=1, max_value=100, value=1, step=1)

    allow_court_gaps = st.checkbox("途中ラウンドの空きコートを許容", value=False)

    html_passcode = st.text_input("HTML簡易ロック用パスコード（任意）", value="", type="password")

    run = st.button("生成", type="primary", use_container_width=True)


def _read_bytes(p: Path) -> bytes:
    return p.read_bytes()


def _set_last_outputs(*, excel_name: str, excel_bytes: bytes, html_name: str | None, html_bytes: bytes | None) -> None:
    st.session_state.last_excel_name = excel_name
    st.session_state.last_excel_bytes = excel_bytes
    st.session_state.last_html_name = html_name
    st.session_state.last_html_bytes = html_bytes


if run:
    if uploaded is None:
        st.error("チーム一覧Excelをアップロードしてください")
        st.stop()

    with st.status("生成中…", expanded=True) as status:
        try:
            with tempfile.TemporaryDirectory() as td:
                tmp_dir = Path(td)
                input_path = tmp_dir / "teams.xlsx"
                input_path.write_bytes(uploaded.getvalue())

                # Place outputs in temp directory; scheduler will stamp names.
                output_base = tmp_dir / "schedule.xlsx"

                # Run the existing pipeline.
                scheduler.generate_schedule(
                    input_file=str(input_path),
                    output_file=str(output_base),
                    num_rounds=int(num_rounds),
                    courts=int(courts),
                    diversity_attempts=int(diversity_attempts),
                    graph_mode=True,
                    allow_court_gaps=bool(allow_court_gaps),
                    max_consecutive=int(max_consecutive),
                    matches_per_team=int(matches_per_team),
                    html_passcode=str(html_passcode),
                    start_time=str(start_time),
                    round_minutes=int(round_minutes),
                )

                # Find generated files (timestamped).
                xlsx_files = sorted(tmp_dir.glob("schedule_*.xlsx"))
                html_files = sorted(tmp_dir.glob("schedule_*.html"))

                if not xlsx_files:
                    raise RuntimeError("Excel出力が見つかりませんでした")

                xlsx_path = xlsx_files[-1]
                html_path = html_files[-1] if html_files else None

                excel_bytes = _read_bytes(xlsx_path)
                html_bytes = _read_bytes(html_path) if html_path else b""

            _set_last_outputs(
                excel_name=xlsx_path.name,
                excel_bytes=excel_bytes,
                html_name=(html_path.name if html_path is not None else None),
                html_bytes=(html_bytes if html_path is not None else None),
            )

            status.update(label="生成完了", state="complete", expanded=False)

            st.success("生成しました。右側のダウンロードからExcel/HTMLを取得できます。")

        except Exception as e:
            status.update(label="失敗", state="error")
            st.exception(e)


if "last_excel_bytes" in st.session_state:
    st.markdown("### ダウンロード")
    st.caption("生成結果はサーバーにファイルとしては残さず、このブラウザのセッション内メモリで保持します（再読み込みすると消えます）。")
    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "Excelをダウンロード",
            data=st.session_state.last_excel_bytes,
            file_name=st.session_state.last_excel_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col2:
        if st.session_state.get("last_html_bytes") and st.session_state.get("last_html_name"):
            st.download_button(
                "HTMLをダウンロード",
                data=st.session_state.last_html_bytes,
                file_name=st.session_state.last_html_name,
                mime="text/html",
                use_container_width=True,
            )
        else:
            st.info("HTML出力はありませんでした")

    st.divider()
    st.caption("注意: HTMLの簡易ロックは暗号化ではありません。配布先は限定してください。")
