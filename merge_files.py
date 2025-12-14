import os
from pathlib import Path
from pypdf import PdfWriter
from docx2pdf import convert

def merge_docs_and_pdfs(source_folder, output_filename="merged_output.pdf"):
    """
    掃描資料夾，將所有 .docx 和 .pdf 檔案合併為一份 PDF。
    
    Args:
        source_folder (str): 包含檔案的資料夾路徑
        output_filename (str): 輸出的檔案名稱
    """
    
    # 初始化 PDF 合併器
    merger = PdfWriter()
    
    # 取得資料夾路徑物件
    folder_path = Path(source_folder)
    
    # 獲取所有檔案並排序 (確保合併順序是依照檔名，例如 1.docx, 2.pdf...)
    files = sorted([f for f in folder_path.iterdir() if f.is_file()])
    
    temp_pdfs = [] # 用來記錄臨時產生的 PDF，以便稍後刪除

    print(f"📂 開始掃描資料夾: {source_folder}")

    try:
        for file_path in files:
            # 忽略隱藏檔案 (如 macOS 的 .DS_Store) 或暫存檔
            if file_path.name.startswith('~') or file_path.name.startswith('.'):
                continue

            # 處理 Word 檔 (.docx)
            if file_path.suffix.lower() == '.docx':
                print(f"🔄 正在轉換 Word 檔: {file_path.name} ...")
                
                # 定義臨時 PDF 路徑
                temp_pdf_path = file_path.with_suffix('.temp.pdf')
                
                # 執行轉換
                try:
                    convert(str(file_path), str(temp_pdf_path))
                    merger.append(str(temp_pdf_path))
                    temp_pdfs.append(temp_pdf_path) # 加入待刪除清單
                    print(f"✅ 已加入: {file_path.name} (已轉為 PDF)")
                except Exception as e:
                    print(f"❌ 轉換失敗 {file_path.name}: {e}")

            # 處理 PDF 檔
            elif file_path.suffix.lower() == '.pdf':
                # 避免將輸出的檔案自己也合併進去 (如果它已存在)
                if file_path.name == output_filename:
                    continue
                    
                print(f"📄 讀取 PDF 檔: {file_path.name}")
                merger.append(str(file_path))
                print(f"✅ 已加入: {file_path.name}")

        # 輸出最終檔案
        output_path = folder_path / output_filename
        print(f"💾 正在寫入最終檔案: {output_path} ...")
        merger.write(str(output_path))
        merger.close()
        print(f"🎉 合併完成！檔案位於: {output_path}")

    except Exception as e:
        print(f"💥 發生嚴重錯誤: {e}")
    
    finally:
        # 清理臨時檔案
        if temp_pdfs:
            print("🧹 正在清理臨時檔案...")
            for temp in temp_pdfs:
                try:
                    os.remove(temp)
                except OSError:
                    pass
            print("✨ 清理完畢")

# ==========================================
# 使用設定區
# ==========================================
if __name__ == "__main__":
    # 設定您的資料夾路徑 (請修改這裡)
    # 例如 Windows: r"C:\Users\Dennis\Documents\ProjectA"
    # 例如 Mac: "/Users/Dennis/Documents/ProjectA"
    
    TARGET_FOLDER = r"./my_documents"  # 預設為程式所在的 my_documents 資料夾
    OUTPUT_NAME = "Final_Report_2025.pdf"

    # 如果資料夾不存在，建立一個範例讓使用者知道放哪裡
    if not os.path.exists(TARGET_FOLDER):
        os.makedirs(TARGET_FOLDER)
        print(f"⚠️ 資料夾 '{TARGET_FOLDER}' 不存在，已為您建立。請將 PDF/Word 檔案放入其中後再次執行。")
    else:
        merge_docs_and_pdfs(TARGET_FOLDER, OUTPUT_NAME)
