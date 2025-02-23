import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

def read_xls_column_to_list(file_path):
    # 讀取 xlsm 檔案
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    
    # 讀取名叫 "美股" 的工作表
    df = pd.read_excel(xls, sheet_name="美股")
    
    # 取得第一欄 (Column 1) 的資料並轉換成 list
    column_1_list = df.iloc[:, 0].tolist()  # iloc[:, 0] 取出第一欄
    
    return column_1_list

def process_files(base_path, file_list, output_file):
    all_data = []
    for file_name in file_list:
        file_path = os.path.join(base_path, f"{file_name}.xlsm")
        if os.path.exists(file_path):
            print(f"Processing file: {file_path}")
            # 讀取對應的 xlsm 檔案
            xls = pd.ExcelFile(file_path, engine='openpyxl')
            df = pd.read_excel(xls, sheet_name="美股", header=None)  # 讀取名叫 "美股" 的工作表
            
            # 取得特定欄位的資料
            data = {
                '代號': file_name,  # 加入 file_name
                'ROE': df.iloc[12, 21],  # V13
                '貴價': df.iloc[6, 10],    # K7
                '淑價': df.iloc[4, 10],    # K5
                '現價': df.iloc[2, 10],    # K3
                '預期報酬': df.iloc[3, 10],    # K4
                '財報': df.iloc[23, 0],   # A24
                '檔案路徑': file_path   # 加入 file_path
            }
            all_data.append(data)
        else:
            print(f"File not found: {file_path}")
    
    # 將所有資料合併到一個 DataFrame
    combined_data = pd.DataFrame(all_data)
    
    # 將資料寫入以當日日期為檔名的 xlsm 檔案
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        combined_data.to_excel(writer, sheet_name='美股', index=False)

def main():
    base_path = r"D:\work\me\what\company\system\資訊處理循環\tools\盈再表"  # 替換成實際的檔案路徑
    a_file_path = os.path.join(base_path, 'a.xlsm')
    
    # 生成以當日日期為檔名的 xlsm 檔案
    today_date = datetime.now().strftime("%Y%m%d")
    output_file = os.path.join(base_path, f'{today_date}.xlsx')
    
    # 呼叫函數，讀取 a.xlsm 檔案並將第一欄資料存入 list
    file_list = read_xls_column_to_list(a_file_path)
    
    # 處理對應的 xlsm 檔案並將資料寫入以當日日期為檔名的 xlsm 檔案
    process_files(base_path, file_list, output_file)

# 檢查是否直接執行此檔案
if __name__ == "__main__":
    main()