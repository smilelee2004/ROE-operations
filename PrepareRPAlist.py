import pandas as pd
from openpyxl import load_workbook, Workbook
import os

def create_empty_excel(file_path):
    wb = Workbook()
    wb.save(file_path)

def copy_column_a_to_b(a_file, b_file):
    # 檢查 b_file 是否存在，如果不存在則創建一個新的檔案
    if not os.path.exists(b_file):
        create_empty_excel(b_file)
    
    # 讀取 a.xlsx 的第一欄
    df_a = pd.read_excel(a_file, usecols=[0], header=None)
    
    # 將第一欄資料寫入 b.xlsx 的第一欄
    with pd.ExcelWriter(b_file, engine='openpyxl', mode='w') as writer:
        # 確保至少有一個可見的工作表
        if not writer.book.sheetnames:
            writer.book.create_sheet("Sheet1")
        # 獲取第一個工作表的名稱
        sheet_name = writer.book.sheetnames[0]
        df_a.to_excel(writer, sheet_name=sheet_name, index=False, header=False, startcol=0)

def copy_b_to_k(b_file):
    # 讀取 b.xlsx
    wb = load_workbook(b_file)
    ws = wb[wb.sheetnames[0]]  # 使用第一個工作表
    ws2 = wb.create_sheet("Sheet2")  # 創建第二個工作表
    all_data = []

    data = {
        'StockID': "StockID",
        'QuarterlyIncomeStatement': f"QuarterlyIncomeStatement",
        'QuarterlyBalanceSheet': f"QuarterlyBalanceSheet",
        'AnnualIncomeStatement': f"AnnualIncomeStatement",
        'AnnualBalanceSheet': f"AnnualBalanceSheet",
        'CompanyNameProfile': f"CompanyNameProfile",
        'AnnualCashflowStatement': f"AnnualCashflowStatement",
        'HistoricalStockPrice': f"HistoricalStockPrice",
    }
    all_data.append(data)
    # 取得第一欄的最大行數
    max_row = ws.max_row
    
    # 複製 1B 到 1K 的公式到 2B 到 2K，直到第一欄的每個有值的欄位
    for row in range(2, max_row + 1):
        # 取得特定欄位的資料
        testString = f"https://www.marketwatch.com/investing/Stock/{ws.cell(row, 1).value}/financials/income/quarter"
        data = {
            'StockID': ws.cell(row, 1).value,
            'QuarterlyIncomeStatement': f"https://www.marketwatch.com/investing/Stock/{ws.cell(row, 1).value}/financials/income/quarter",
            'QuarterlyBalanceSheet': f"https://www.marketwatch.com/investing/Stock/{ws.cell(row, 1).value}/financials/balance-sheet/quarter",
            'AnnualIncomeStatement': f"https://www.marketwatch.com/investing/Stock/{ws.cell(row, 1).value}/financials",
            'AnnualBalanceSheet': f"https://www.marketwatch.com/investing/Stock/{ws.cell(row, 1).value}/financials/balance-sheet",
            'CompanyNameProfile': f"https://www.marketwatch.com/investing/stock/{ws.cell(row, 1).value}/company-profile",
            'AnnualCashflowStatement': f"https://www.marketwatch.com/investing/Stock/{ws.cell(row, 1).value}/financials/cash-flow",
            'HistoricalStockPrice': f"https://finance.yahoo.com/quote/{ws.cell(row, 1).value}/history?period1=573436800&period2=1740960000&interval=1mo&filter=history&frequency=1mo",
        }
        all_data.append(data)

    # 將 all_data 寫入 ws2 工作表
    for data in all_data:
        ws2.append(list(data.values()))

    # 刪除 Sheet1 工作表
    wb.remove(ws)
    
    # 儲存 b.xlsx
    wb.save(b_file)

def main():
    a_file = r"D:\work\me\what\company\system\資訊處理循環\tools\RPA\a.xlsx"
    b_file = r"D:\work\me\what\company\system\資訊處理循環\tools\RPA\RPA下載清單.xlsx"
    
    # 複製 a.xlsx 的第一欄到 b.xlsx 的第一欄
    copy_column_a_to_b(a_file, b_file)
    
    # 複製 b.xlsx 中 1B 到 1K 的資料到 2B 到 2K，直到第一欄的每個有值的欄位
    copy_b_to_k(b_file)

if __name__ == "__main__":
    main()