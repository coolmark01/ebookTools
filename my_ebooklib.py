import os
import xlwt
from ebooklib import epub

def scan_epub_files(base_folder, excel_file):
    # 創建 Excel 檔案
    wb = xlwt.Workbook()
    ws = wb.add_sheet("書籍資料")

    # 設置欄位名稱
    ws.write(0, 0, "書名")
    ws.write(0, 1, "作者")
    ws.write(0, 2, "出版社")
    ws.write(0, 3, "出版日期")
    ws.write(0, 4, "語言")

    row_num = 1  # 從第二行開始寫入資料
    for foldername, subfolders, filenames in os.walk(base_folder):
        for filename in filenames:
            # 獲取檔案的名稱和副檔名
            file_name, file_ext = os.path.splitext(filename)
            file_ext = file_ext.lower()

            # 例外排除指定副檔名
            if file_ext != '.epub':
                continue

            # 獲取完整檔案路徑
            file_path = os.path.join(foldername, filename)

            # 解析 ePub 檔案
            try:
                book = epub.read_epub(file_path)
            except Exception as e:
                print("無法讀取 ePub 檔案:", file_path)
                continue
            
            # 獲取 Metadata
            title = book.get_metadata('DC', 'title')
            authors = book.get_metadata('DC', 'creator')
            publishers = book.get_metadata('DC', 'publisher')
            dates = book.get_metadata('DC', 'date')
            languages = book.get_metadata('DC', 'language')

            # 將資訊寫入 Excel
            ws.write(row_num, 0, title[0][0])
            ws.write(row_num, 1, authors[0][0])
            ws.write(row_num, 2, publishers[0][0] if publishers else "")
            ws.write(row_num, 3, dates[0][0] if dates else "")
            ws.write(row_num, 4, languages[0][0] if languages else "")
            row_num += 1

    # 儲存 Excel 檔案
    wb.save(excel_file)

if __name__ == "__main__":
    # 獲取程式碼所在的資料夾路徑
    current_folder = os.path.dirname(os.path.abspath(__file__))

    # 設置 Excel 檔案名稱
    excel_file = "ebook_metadata.xls"  # 注意這裡使用 .xls 格式

    scan_epub_files(current_folder, excel_file)
    print("掃描完成，Excel檔案已儲存為：", excel_file)
