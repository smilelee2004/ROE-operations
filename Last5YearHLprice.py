import pandas as pd
import yfinance as yf
import csv

def read_xls_column_to_list(file_path):
    # 讀取 xls 檔案
    xls = pd.ExcelFile(file_path)
    
    # 讀取 sheet 1 的資料
    df = pd.read_excel(xls, sheet_name=0)  # 0 表示第一個 sheet
    
    # 取得第一欄 (Column 1) 的資料並轉換成 list
    column_1_list = df.iloc[:, 0].tolist()  # iloc[:, 0] 取出第一欄
    

    # Add additional comments to test GitHub.
    return column_1_list



def main():

    # 建立一個空的二維列表
    myHighestLowestList = []
    myHighestLowestList.append(["Symbol", "High", "LOW"])
    # myHighestLowestList = [{"Symbol","High","LOW"}]

    # 呼叫函數，讀取檔案並將第一欄資料存入 list
    file_path = "D:\\work\\TestFolder\\"  # 替換成實際的檔案路徑
    file_name = 'MyStockList2.xls'
    column_data_list = read_xls_column_to_list(file_path+file_name)

    # 印出 list
    #print(column_data_list)
    for item in column_data_list:
        #print(item)

        # 判斷是否為字串
        if isinstance(item, str):
            # 美股
            # modified_string = item.replace(".", "-")
            ko = yf.Ticker(item.replace(".", "-"))
        else:
            # 台股
            ko = yf.Ticker(str(item)+".TW")       
      
 
        quarterly_financials = ko.quarterly_financials

        # 抓取財務報表
        #print("財務報表:")
        # financial_data.to_csv("ko_financial.csv")
        # balance_sheet.to_csv("ko_balancesheet.csv")
        #cashflow_data.to_csv("ko_cashflow.csv")
        #quarterly_financials.to_csv("ko_quarterly_financial.csv")

        #print(ko.financials)  # 獲取收入報表
        #print("\n資產負債表:")
        #print(ko.balance_sheet)  # 資產負債表
        #print("\n現金流報表:")
        #print(ko.cashflow)  # 現金流報表

        # 抓取歷史股價數據（過去一年）
        history_data = ko.history(period="5y")  # 抓取過去五年的數據
        if history_data.empty:
            history_data = ko.history(period="max")  # 抓取過去五年的數據
            if history_data.empty:
                print(f"{item }獲取數據失敗或沒有資料。")
                myHighestLowestList.append([item, "na", "na"])
            
        if history_data.empty == False:
            # 找出最高價與最低價
            highest_price = history_data['Close'].max()
            lowest_price = history_data['Close'].min()
            # 動態添加行
            myHighestLowestList.append([item, highest_price, lowest_price])
        
            # 保存歷史數據到 CSV 文件
            #filename = f"{item}_price_history.csv"
            #history_data.to_csv(filename)
            print(f"\n{item}, High {highest_price}, Low {lowest_price}")

    # 將資料寫入 CSV 檔案
    with open(file_path + 'HighLow'+ file_name.removesuffix(".xls")  + '.csv', mode='w', newline='') as file:
        writer = csv.writer(file)
        # 寫入每一行
        writer.writerows(myHighestLowestList)

    csvdata = pd.read_csv(file_path + 'HighLow'+ file_name.removesuffix(".xls")  + '.csv')
    csvdata.to_excel( file_path + 'HighLow'+ file_name + "x", index=False, engine='openpyxl')

# 檢查是否直接執行此檔案
if __name__ == "__main__":
    main()
