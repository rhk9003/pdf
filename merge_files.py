import io
import pandas as pd
import streamlit as st
from pypdf import PdfWriter

st.set_page_config(page_title="檔案合併工具箱", page_icon="🧰", layout="centered")

st.title("🧰 檔案合併工具箱")
st.caption("Excel：多工作表直向堆疊成單一工作表｜PDF：多檔合併成一份 PDF")

tool = st.sidebar.radio(
    "選擇功能",
    ["Excel 多工作表 → 單一工作表（直向堆疊）", "PDF 多檔 → 合併成一份"],
)

# =========================================================
# Excel：多工作表直向堆疊
# =========================================================
def read_all_sheets(xlsx_bytes: bytes) -> dict[str, pd.DataFrame]:
    xl = pd.ExcelFile(io.BytesIO(xlsx_bytes), engine="openpyxl")
    sheets: dict[str, pd.DataFrame] = {}
    for name in xl.sheet_names:
        # dtype=str：避免多表型別混雜導致 concat 不穩；若你想保留原型別可移除 dtype=str
        df = xl.parse(sheet_name=name, dtype=str)
        sheets[name] = df
    return sheets

def stack_sheets(
    sheets: dict[str, pd.DataFrame],
    add_sheet_col: bool = True,
    sheet_col_name: str = "__sheet__",
    keep_all_cols: bool = True,
) -> pd.DataFrame:
    frames = []
    for sname, df in sheets.items():
        if df is None or df.empty:
            continue
        df2 = df.copy()
        if add_sheet_col:
            df2.insert(0, sheet_col_name, sname)
        frames.append(df2)

    if not frames:
        return pd.DataFrame()

    if keep_all_cols:
        return pd.concat(frames, ignore_index=True, sort=False)

    # 欄位取交集（可用但較不建議）
    common = set(frames[0].columns)
    for f in frames[1:]:
        common &= set(f.columns)
    common_cols = [c for c in frames[0].columns if c in common]
    return pd.concat([f[common_cols] for f in frames], ignore_index=True)

def write_single_sheet_xlsx(df: pd.DataFrame, sheet_name: str) -> bytes:
    out = io.BytesIO()
    safe_name = (sheet_name or "merged")[:31]  # Excel sheet name <= 31 chars
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=safe_name)
    return out.getvalue()

if tool.startswith("Excel"):
    st.subheader("📎 Excel 多工作表 → 單一工作表（直向堆疊）")

    uploaded = st.file_uploader("上傳 Excel 檔（.xlsx）", type=["xlsx"], key="excel_uploader")

    with st.expander("設定", expanded=True):
        add_sheet_col = st.checkbox("新增來源工作表欄位", value=True, key="excel_add_sheet_col")
        sheet_col_name = st.text_input("來源工作表欄位名稱", value="__sheet__", key="excel_sheet_col_name")
        keep_all_cols = st.checkbox("欄位取聯集（不同工作表欄位不同也保留）", value=True, key="excel_keep_all_cols")
        output_sheet_name = st.text_input("輸出工作表名稱", value="merged", key="excel_output_sheet_name")

    if uploaded:
        try:
            raw = uploaded.getvalue()
            sheets = read_all_sheets(raw)

            st.success(f"讀取成功：{len(sheets)} 個工作表")
            st.write("工作表：", ", ".join(sheets.keys()))

            merged_df = stack_sheets(
                sheets,
                add_sheet_col=add_sheet_col,
                sheet_col_name=sheet_col_name,
                keep_all_cols=keep_all_cols,
            )

            st.subheader("預覽（前 200 列）")
            st.dataframe(merged_df.head(200), use_container_width=True)

            output_bytes = write_single_sheet_xlsx(merged_df, output_sheet_name)

            st.download_button(
                label="⬇️ 下載合併後 Excel（單一工作表）",
                data=output_bytes,
                file_name=f"{uploaded.name.rsplit('.', 1)[0]}__stacked.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error("處理失敗：可能不是標準 .xlsx 或內容結構不支援。")
            st.exception(e)
    else:
        st.info("請先上傳一個含多工作表的 .xlsx 檔案。")

# =========================================================
# PDF：多檔合併
# =========================================================
else:
    st.subheader("📄 PDF 多檔 → 合併成一份")

    st.markdown(
        """
**注意**：此工具僅處理 PDF 檔合併。  
若你需要合併 Word，通常要在本機環境或額外轉檔流程。
"""
    )

    uploaded_files = st.file_uploader(
        "請選擇要合併的 PDF 檔案（可多選）",
        type=["pdf"],
        accept_multiple_files=True,
        key="pdf_uploader",
    )

    if uploaded_files:
        st.write(f"✅ 已選擇 {len(uploaded_files)} 個檔案")

        with st.expander("查看檔案清單與順序"):
            for i, f in enumerate(uploaded_files):
                st.text(f"{i+1}. {f.name}")

        output_name = st.text_input("輸出檔名（.pdf 會自動補上）", value="merged_document", key="pdf_output_name")

        if st.button("開始合併 PDF", key="pdf_merge_btn"):
            try:
                merger = PdfWriter()
                progress = st.progress(0)

                for idx, pdf_file in enumerate(uploaded_files):
                    merger.append(pdf_file)
                    progress.progress((idx + 1) / len(uploaded_files))

                buf = io.BytesIO()
                merger.write(buf)
                merger.close()
                buf.seek(0)

                st.success("🎉 合併完成！請下載。")

                filename = output_name.strip() or "merged_document"
                if not filename.lower().endswith(".pdf"):
                    filename += ".pdf"

                st.download_button(
                    label="📥 下載合併後的 PDF",
                    data=buf,
                    file_name=filename,
                    mime="application/pdf",
                )

            except Exception as e:
                st.error(f"發生錯誤：{e}")
    else:
        st.info("請上傳 PDF 檔案以開始使用。")
