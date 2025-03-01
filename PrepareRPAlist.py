import pandas as pd
from openpyxl import load_workbook

def copy_column_a_to_b(a_file, b_file):
    # 讀取 a.xlsx 的第一欄
    df_a = pd.read_excel(a_file, usecols=[0], header=None)
    
    # 將第一欄資料寫入 b.xlsx 的第一欄
    with pd.ExcelWriter(b_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
        # 獲取第一個工作表的名稱
        sheet_name = writer.book.sheetnames[0]
        df_a.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startcol=0)

def copy_b_to_k(b_file):
    # 讀取 b.xlsx
    wb = load_workbook(b_file)
    ws = wb[wb.sheetnames[0]]  # 使用第一個工作表
    ws2 = wb[wb.sheetnames[1]]  # 使用第二個工作表

    # 取得第一欄的最大行數
    max_row = ws.max_row
    
    # 複製 1B 到 1K 的公式到 2B 到 2K，直到第一欄的每個有值的欄位
    for row in range(2, max_row):
        for col in range(2, 12):  # B 到 K 對應的列號是 2 到 11
            ws.cell(row=row+1, column=col).value = ws.cell(row=2, column=col).value
            ws2.cell(row=row+1, column=col).value = ws2.cell(row=2, column=col).value
            if ws.cell(row=2, column=col).data_type == 'f':  # 檢查是否為公式
                ws.cell(row=row+1, column=col).value = ws.cell(row=2, column=col).value
                ws2.cell(row=row+1, column=col).value = ws2.cell(row=2, column=col).value
    
    # 儲存 b.xlsx
    wb.save(b_file)

def main():
    a_file = r"D:\work\me\what\company\system\資訊處理循環\tools\RPA\a.xlsx"
    b_file = r"D:\work\me\what\company\system\資訊處理循環\tools\RPA\RPA下載清單 sample .xlsx"
    
    # 複製 a.xlsx 的第一欄到 b.xlsx 的第一欄
    copy_column_a_to_b(a_file, b_file)
    
    # 複製 b.xlsx 中 1B 到 1K 的資料到 2B 到 2K，直到第一欄的每個有值的欄位
    copy_b_to_k(b_file)

if __name__ == "__main__":
    main()