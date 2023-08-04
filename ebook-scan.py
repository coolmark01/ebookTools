import os
import xlwt

def scan_files(base_folder, excel_file):
    # 創建 Excel 檔案
    wb = xlwt.Workbook()
    ws = wb.add_sheet("書籍資料")

    # 設置欄位名稱
    ws.write(0, 0, "書名")
    ws.write(0, 1, "格式")

    row_num = 1  # 從第二行開始寫入資料
    for foldername, subfolders, filenames in os.walk(base_folder):
        for filename in filenames:
            # 獲取檔案的名稱和副檔名
            file_name, file_ext = os.path.splitext(filename)
            file_ext = file_ext.lower()

            # 例外排除指定副檔名
            if file_ext in ['.jpg', '.png', '.bmp', '.py', '.xlsx', '.xls']:
                continue

            # 將檔案名稱和副檔名寫入 Excel
            ws.write(row_num, 0, file_name)
            ws.write(row_num, 1, file_ext)
            row_num += 1

    # 儲存 Excel 檔案
    wb.save(excel_file)

if __name__ == "__main__":
    # 獲取程式碼所在的資料夾路徑
    current_folder = os.path.dirname(os.path.abspath(__file__))

    # 設置 Excel 檔案名稱
    excel_file = "ebook_resources.xls"  # 注意這裡使用 .xls 格式

    scan_files(current_folder, excel_file)
    print("掃描完成，Excel檔案已儲存為：", excel_file)
