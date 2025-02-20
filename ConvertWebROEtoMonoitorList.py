import pandas as pd
import csv
import yfinance as yf

def read_xls_column_to_list(file_path):
    # 讀取 xls 檔案
    xls = pd.ExcelFile(file_path)
    
    # 讀取 sheet 1 的資料
    df = pd.read_excel(xls, sheet_name=0)  # 0 表示第一個 sheet
    
    # 取得第一欄 (Column 1) 的資料並轉換成 list
    column_1_list = df.iloc[:, 0].tolist()  # iloc[:, 0] 取出第一欄
    
    return column_1_list

def getHighLowPrice():

    # 建立一個空的二維列表
    myHighestLowestList = []
    myHighestLowestList.append(["Symbol", "High", "LOW"])
    # myHighestLowestList = [{"Symbol","High","LOW"}]

    # 呼叫函數，讀取檔案並將第一欄資料存入 list
    file_path = "D:\\work\\TestFolder\\"  # 替換成實際的檔案路徑
    file_name = 'MyStockList.xls'
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
                myHighestLowestList.append([item, "  ", "  "])
            
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



def main():

    # 呼叫函數, 取得過去5年的最高最低價格
    # getHighLowPrice()

    # 呼叫函數，讀取檔案並將第一欄資料存入 list
    # file_path = 'D:\work\TestFolder\myStockList2.xls'  # 替換成實際的檔案路徑
    # column_data_list = read_xls_column_to_list(file_path)
    # 讀取 c.xls 和 d.xls 檔案
    c_df = pd.read_excel('D:\\work\\TestFolder\\MyStockList.xls')
    d_df = pd.read_excel('D:\\work\\TestFolder\\uslist.xlsx')
    e_df = pd.read_excel('D:\\work\\TestFolder\\HighLowMyStockList.xlsx')


    # 提取兩個檔案的第一行 (header) 作為 column 標題
    c_columns = c_df.columns

    d_columns = d_df.columns

    # 找出 d.xls 中 row 1 不存在於 c.xls 的 column
    columns_to_remove = [col for col in d_columns if col not in c_columns]

    # 移除這些 column
    d_df_cleaned = d_df.drop(columns=columns_to_remove)



    # 根據 column 1 的內容來篩選 d.xls 的行，移除不在 c.xls 中 column 1 的行
    filtered_d_df = d_df_cleaned[d_df_cleaned[d_df.columns[0]].isin(c_df[c_df.columns[0]])]


    #在合併前，將兩個 DataFrame 的 Symbol 欄位轉換為相同的數據類型。例如，可以將它們都轉換為字串類型
    e_df['Symbol'] = e_df['Symbol'].astype(str)
    filtered_d_df.loc[:, 'Symbol'] = filtered_d_df['Symbol'].astype(str)
    # 比對 A.xls 和 B.xls 的 column A
    # 假設 A 列是唯一識別碼，可以作為合併的關鍵
    merged_df = pd.merge(e_df,filtered_d_df,  on='Symbol', how='left')
   

    # 如果 B.xls 中的相應資料存在，將其附加到 A.xls 的對應行上
    # 注意：`on='A'` 是合併的關鍵，`how='left'` 表示保留 A.xls 的所有行
    # 將結果寫回 c.xls

    merged_df.to_excel('D:\\work\\TestFolder\\d_cleaned.xlsx', index=False)
    print("完成")


# 檢查是否直接執行此檔案
if __name__ == "__main__":
    main()
