import streamlit as st
from pypdf import PdfWriter
import io

# 設定頁面資訊
st.set_page_config(page_title="PDF 合併工具", page_icon="📄")

st.title("📄 PDF 文件合併工具")
st.markdown("""
**注意**：由於雲端環境限制 (無 Microsoft Word)，此線上版本僅支援 **PDF 檔案** 的合併。
若需合併 Word 檔，請在您本機電腦執行。
""")

# 1. 檔案上傳區
uploaded_files = st.file_uploader(
    "請選擇要合併的 PDF 檔案 (可多選，並拖拉排序)", 
    type="pdf", 
    accept_multiple_files=True
)

if uploaded_files:
    st.write(f"✅ 已選擇 {len(uploaded_files)} 個檔案")
    
    # 讓使用者確認順序 (Streamlit 上傳後通常依照檔名，但使用者可透過重新命名控制)
    # 這裡我們簡單列出清單
    with st.expander("查看檔案清單與順序"):
        for i, file in enumerate(uploaded_files):
            st.text(f"{i+1}. {file.name}")

    # 2. 合併按鈕
    if st.button("開始合併 PDF"):
        try:
            merger = PdfWriter()
            
            # 進度條
            progress_bar = st.progress(0)
            
            for index, pdf_file in enumerate(uploaded_files):
                merger.append(pdf_file)
                # 更新進度條
                progress_bar.progress((index + 1) / len(uploaded_files))
            
            # 將合併結果寫入記憶體 (不存硬碟，適合雲端)
            output_buffer = io.BytesIO()
            merger.write(output_buffer)
            merger.close()
            
            # 重置游標位置到開頭，以便讀取
            output_buffer.seek(0)
            
            st.success("🎉 合併完成！請點擊下方按鈕下載。")
            
            # 3. 下載按鈕
            st.download_button(
                label="📥 下載合併後的 PDF",
                data=output_buffer,
                file_name="merged_document.pdf",
                mime="application/pdf"
            )
            
        except Exception as e:
            st.error(f"發生錯誤：{e}")

else:
    st.info("請上傳檔案以開始使用。")
