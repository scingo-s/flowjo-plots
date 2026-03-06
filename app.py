"""FlowJo Plot 画像差替ツール - Streamlit Web アプリ."""

import streamlit as st

from replacer import process

st.set_page_config(page_title="FlowJo Plot 画像差替", layout="wide")


def check_password():
    """パスワード認証ゲート."""
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False

    if st.session_state.authenticated:
        return True

    password = st.text_input("パスワードを入力してください", type="password")
    if password:
        if password == st.secrets.get("password", ""):
            st.session_state.authenticated = True
            st.rerun()
        else:
            st.error("パスワードが正しくありません")
    return False


def main():
    st.title("FlowJo Plot 画像差替ツール")
    st.caption("FlowJo バッチ出力 PDF のプロット画像を Excel に自動貼付します")

    if not check_password():
        return

    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader(
            "FlowJo バッチ出力 PDF",
            type=["pdf"],
            help="14サンプル x 8プロットのグリッド PDF",
        )
    with col2:
        xlsx_file = st.file_uploader(
            "Summary Excel (.xlsx)",
            type=["xlsx"],
            help="画像を差し替える先の Excel ファイル",
        )

    if pdf_file is None or xlsx_file is None:
        st.info("PDF と Excel の両方をアップロードしてください")
        return

    with st.spinner("画像を差し替え中...（数十秒かかります）"):
        pdf_bytes = pdf_file.read()
        xlsx_bytes = xlsx_file.read()
        result = process(pdf_bytes, xlsx_bytes)

    st.success("差し替え完了 (112枚)")

    st.download_button(
        "差し替え済み Excel をダウンロード",
        data=result,
        file_name=f"Summary_with_plots.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


if __name__ == "__main__":
    main()
