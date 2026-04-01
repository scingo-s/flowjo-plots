"""FlowJo Plot 画像差替ツール - Streamlit Web アプリ."""

import streamlit as st

from replacer import (
    create_preview_grid,
    extract_plots_from_pdf,
    insert_images_to_xlsx,
    validate_excel,
    validate_extracted_images,
    validate_pdf,
)

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


def _reset_state():
    """処理ステートをリセットする."""
    for key in ("step", "plots", "num_rows", "num_cols", "preview_grid",
                "xlsx_bytes", "result", "img_validation",
                "pdf_warnings", "xlsx_warnings"):
        st.session_state.pop(key, None)


def show_upload_step():
    """Step 1: ファイルアップロード + 画像抽出."""
    col1, col2 = st.columns(2)
    with col1:
        pdf_file = st.file_uploader(
            "FlowJo バッチ出力 PDF",
            type=["pdf"],
            help="サンプル x プロットのグリッド PDF（サンプル数は自動検出）",
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

    if st.button("画像を抽出してプレビュー", type="primary"):
        try:
            pdf_bytes = pdf_file.read()
            xlsx_bytes = xlsx_file.read()

            # バリデーション
            pdf_warnings = validate_pdf(pdf_bytes)
            xlsx_warnings = validate_excel(xlsx_bytes)

            # 画像抽出
            with st.spinner("PDF から画像を抽出中..."):
                plots, num_rows, num_cols = extract_plots_from_pdf(pdf_bytes)
                img_validation = validate_extracted_images(plots)
                preview_grid = create_preview_grid(plots, cols=num_cols, rows=num_rows)

            # ステートに保存して次のステップへ
            st.session_state.step = "preview"
            st.session_state.plots = plots
            st.session_state.num_rows = num_rows
            st.session_state.num_cols = num_cols
            st.session_state.preview_grid = preview_grid
            st.session_state.xlsx_bytes = xlsx_bytes
            st.session_state.img_validation = img_validation
            st.session_state.pdf_warnings = pdf_warnings
            st.session_state.xlsx_warnings = xlsx_warnings
            st.rerun()

        except Exception as e:
            st.error(f"エラー: {e}")
            import traceback
            st.code(traceback.format_exc())


def show_preview_step():
    """Step 2: プレビュー確認 + Excel 挿入."""
    # 警告・情報表示
    for w in st.session_state.get("pdf_warnings", []):
        if w.startswith("info:"):
            st.info(w[5:])
        else:
            st.warning(w)
    for w in st.session_state.get("xlsx_warnings", []):
        st.warning(w)

    # 画像品質チェック結果
    v = st.session_state.img_validation
    blank = v["blank_count"]
    total = v["total"]
    if blank > 0:
        st.warning(f"空白（ほぼ単色）の画像が {blank}/{total} 枚検出されました。"
                   "PDF ファイルが正しいか確認してください。")
    else:
        st.success(f"{total} 枚の画像を正常に抽出しました")

    # プレビュー表示
    nr = st.session_state.get("num_rows", "?")
    nc = st.session_state.get("num_cols", "?")
    st.subheader(f"抽出画像プレビュー（{nr}行 x {nc}列）")
    st.image(st.session_state.preview_grid, use_container_width=True)

    # アクションボタン
    col1, col2 = st.columns(2)
    with col1:
        if st.button("Excel に挿入", type="primary"):
            try:
                with st.spinner("Excel に画像を挿入中..."):
                    result = insert_images_to_xlsx(
                        st.session_state.xlsx_bytes,
                        st.session_state.plots,
                        st.session_state.num_rows,
                        st.session_state.num_cols,
                    )
                st.session_state.step = "done"
                st.session_state.result = result
                st.rerun()
            except Exception as e:
                st.error(f"エラー: {e}")
                import traceback
                st.code(traceback.format_exc())
    with col2:
        if st.button("やり直す"):
            _reset_state()
            st.rerun()


def show_done_step():
    """Step 3: ダウンロード."""
    st.success("差し替え完了!")

    st.download_button(
        "差し替え済み Excel をダウンロード",
        data=st.session_state.result,
        file_name="Summary_with_plots.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )

    if st.button("新しいファイルで処理"):
        _reset_state()
        st.rerun()


def main():
    st.title("FlowJo Plot 画像差替ツール")
    st.caption("FlowJo バッチ出力 PDF のプロット画像を Excel に自動貼付します (v2.1)")

    if not check_password():
        return

    # ステート初期化
    if "step" not in st.session_state:
        st.session_state.step = "upload"

    step = st.session_state.step
    if step == "upload":
        show_upload_step()
    elif step == "preview":
        show_preview_step()
    elif step == "done":
        show_done_step()


if __name__ == "__main__":
    main()
