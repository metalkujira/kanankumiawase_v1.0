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

    uploaded = st.file_uploader("チーム一覧Excel (.xlsx)", type=["xlsx"], key="teams_xlsx") 

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

    diversity_attempts = st.number_input("分散最大化の試行回数（最小100 / 最大300）", min_value=100, max_value=300, value=100, step=1)
    st.caption("注: 300回に近づけるほど改善する可能性はありますが、計算時間が大きく伸びることがあります。")

    allow_court_gaps = st.checkbox("途中ラウンドの空きコートを許容", value=False)

    html_passcode = st.text_input("HTML簡易ロック用パスコード（任意）", value="", type="password")

    html_include_members = st.checkbox(
        "生成HTMLに選手名（氏名）を含める",
        value=False,
        help="ONにすると個人情報が含まれます。公開URLでの配布は避け、配布先を限定してください。",
    )
    if html_include_members:
        st.warning("注意: 選手名（氏名）を含むHTMLを生成します。公開URLでの配布は避け、配布先を限定してください。")
    else:
        st.info("個人情報対策: 生成HTMLは選手名（氏名）を表示しません（ペア名のみ）。")

    run = st.button("生成", type="primary", use_container_width=True)

    st.divider()
    st.header("編集後Excel→HTML再生成")
    st.caption("編集したスケジュールExcel（対戦表/ペア一覧）をアップロードして、HTMLを作り直します。")
    edited_schedule = st.file_uploader("編集後スケジュールExcel (.xlsx)", type=["xlsx"], key="edited_schedule_xlsx")
    regen_include_members = st.checkbox(
        "HTMLに選手名（氏名）を含める",
        value=False,
        help="ONにすると個人情報が含まれます。サーバ側で処理される点に注意してください。",
    )
    if regen_include_members:
        st.warning("注意: 選手名（氏名）を含むHTMLを生成します。公開URLでの配布は避け、配布先を限定してください。")
    else:
        st.info("個人情報対策: 選手名（氏名）を表示しません（ペア名のみ）。")
    regen_passcode = st.text_input("再生成HTMLのパスコード（任意）", value="", type="password", key="regen_html_pass")
    regen = st.button("HTMLを再生成", use_container_width=True)


def _read_bytes(p: Path) -> bytes:
    return p.read_bytes()


def _set_last_outputs(*, excel_name: str, excel_bytes: bytes, html_name: str | None, html_bytes: bytes | None) -> None:
    st.session_state.last_excel_name = excel_name
    st.session_state.last_excel_bytes = excel_bytes
    st.session_state.last_html_name = html_name
    st.session_state.last_html_bytes = html_bytes


def _set_regen_html(*, html_name: str, html_bytes: bytes, include_members: bool) -> None:
    st.session_state.regen_html_name = html_name
    st.session_state.regen_html_bytes = html_bytes
    st.session_state.regen_include_members = include_members


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

                if not xlsx_files:
                    raise RuntimeError("Excel出力が見つかりませんでした")

                xlsx_path = xlsx_files[-1]

                # Streamlitでは、HTMLに選手名を含めるか選択できる。
                safe_html_path = xlsx_path.with_suffix(".html")
                matches, teams, inferred_rounds, inferred_courts = scheduler.load_schedule_from_xlsx(
                    str(xlsx_path),
                    fallback_start_time_hhmm=str(start_time),
                    fallback_round_minutes=int(round_minutes),
                )
                scheduler.write_personal_schedule_html(
                    matches,
                    teams,
                    str(safe_html_path),
                    num_rounds=int(inferred_rounds),
                    courts=int(inferred_courts),
                    html_passcode=(str(html_passcode) or None),
                    start_time_hhmm=str(start_time),
                    round_minutes=int(round_minutes),
                    include_members=bool(html_include_members),
                )

                excel_bytes = _read_bytes(xlsx_path)
                html_bytes = _read_bytes(safe_html_path)

            _set_last_outputs(
                excel_name=xlsx_path.name,
                excel_bytes=excel_bytes,
                html_name=safe_html_path.name,
                html_bytes=html_bytes,
            )

            status.update(label="生成完了", state="complete", expanded=False)

            st.success("生成しました。右側のダウンロードからExcel/HTMLを取得できます。")

        except Exception as e:
            status.update(label="失敗", state="error")
            st.exception(e)


if "edited_schedule_xlsx" in st.session_state and st.session_state.get("edited_schedule_xlsx") is not None:
    # no-op placeholder (Streamlit stores uploader state);
    # actual handling is done on button click below.
    pass


if 'regen' in locals() and regen:
    if edited_schedule is None:
        st.error("編集後スケジュールExcelをアップロードしてください")
        st.stop()

    with st.status("HTML再生成中…", expanded=True) as status:
        try:
            with tempfile.TemporaryDirectory() as td:
                tmp_dir = Path(td)
                input_path = tmp_dir / "edited_schedule.xlsx"
                input_path.write_bytes(edited_schedule.getvalue())

                matches, teams, inferred_rounds, inferred_courts = scheduler.load_schedule_from_xlsx(
                    str(input_path),
                    fallback_start_time_hhmm=str(start_time),
                    fallback_round_minutes=int(round_minutes),
                )

                out_name = Path(edited_schedule.name).with_suffix(".html").name if edited_schedule.name else "schedule.html"
                out_path = tmp_dir / out_name
                scheduler.write_personal_schedule_html(
                    matches,
                    teams,
                    str(out_path),
                    num_rounds=int(inferred_rounds),
                    courts=int(inferred_courts),
                    html_passcode=(str(regen_passcode) or None),
                    start_time_hhmm=str(start_time),
                    round_minutes=int(round_minutes),
                    include_members=bool(regen_include_members),
                )
                html_bytes = _read_bytes(out_path)

            _set_regen_html(html_name=out_name, html_bytes=html_bytes, include_members=bool(regen_include_members))
            status.update(label="再生成完了", state="complete", expanded=False)
            if regen_include_members:
                st.success("HTMLを再生成しました（選手名あり）。下のダウンロードから取得できます。")
            else:
                st.success("HTMLを再生成しました（個人情報なし）。下のダウンロードから取得できます。")
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


if st.session_state.get("regen_html_bytes") and st.session_state.get("regen_html_name"):
    if st.session_state.get("regen_include_members"):
        st.markdown("### 再生成HTML（選手名あり）")
    else:
        st.markdown("### 再生成HTML（個人情報なし）")
    st.download_button(
        "再生成HTMLをダウンロード",
        data=st.session_state.regen_html_bytes,
        file_name=st.session_state.regen_html_name,
        mime="text/html",
        use_container_width=True,
    )
