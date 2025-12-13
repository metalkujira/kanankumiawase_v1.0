from __future__ import annotations

import io
import tempfile
from pathlib import Path

import streamlit as st

from badminton_program import scheduler


st.set_page_config(page_title="Badminton Scheduler", layout="wide")

st.title("Badminton Scheduler")
st.caption("Excelのチームリストをアップロードして、コート数/ラウンド数/各ペア試合数を指定して生成します。")
st.warning(
    "公開運用の注意: アップロードされたExcelはサーバー側で処理されます。個人情報が含まれる場合は匿名化/最小化してからアップロードしてください。",
)

with st.sidebar:
    st.header("設定")

    uploaded = st.file_uploader("チーム一覧Excel (.xlsx)", type=["xlsx"]) 

    courts = st.number_input("コート数", min_value=1, max_value=60, value=15, step=1)
    num_rounds = st.number_input("ラウンド数", min_value=1, max_value=200, value=23, step=1)
    matches_per_team = st.number_input("各ペアの試合数 (0=自動)", min_value=0, max_value=30, value=0, step=1)

    diversity_attempts = st.number_input("分散最大化の試行回数", min_value=1, max_value=50, value=1, step=1)

    allow_court_gaps = st.checkbox("途中ラウンドの空きコートを許容", value=False)

    html_passcode = st.text_input("HTML簡易ロック用パスコード（任意）", value="", type="password")

    run = st.button("生成", type="primary", use_container_width=True)


def _read_bytes(p: Path) -> bytes:
    return p.read_bytes()


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
                    matches_per_team=int(matches_per_team),
                    html_passcode=str(html_passcode),
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

            status.update(label="生成完了", state="complete", expanded=False)

            st.success("生成しました。ダウンロードしてください。")

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "Excelをダウンロード",
                    data=excel_bytes,
                    file_name=xlsx_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with col2:
                if html_path is not None:
                    st.download_button(
                        "HTMLをダウンロード",
                        data=html_bytes,
                        file_name=html_path.name,
                        mime="text/html",
                        use_container_width=True,
                    )
                else:
                    st.info("HTML出力はありませんでした")

            st.divider()
            st.caption("注意: HTMLの簡易ロックは暗号化ではありません。配布先は限定してください。")

        except Exception as e:
            status.update(label="失敗", state="error")
            st.exception(e)
