import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook

def read_xls_column_to_list(file_path):
    # 讀取 xlsm 檔案
    xls = pd.ExcelFile(file_path, engine='openpyxl')
    
    # 讀取第一個工作表
    df = pd.read_excel(xls, sheet_name=0)
    
    # 取得第一欄 (Column 1) 的資料並轉換成 list
    column_1_list = df.iloc[:, 0].tolist()  # iloc[:, 0] 取出第一欄
    
    return column_1_list

def process_files(base_path, file_list, output_file):
    all_data = []
    count = 0
    headers_written = False

    for file_name in file_list:
        file_path = os.path.join(base_path, f"{file_name}.xlsm")
        if os.path.exists(file_path):
            print(f"Processing file: {file_path}")
            try:
                # 讀取對應的 xlsm 檔案
                xls = pd.ExcelFile(file_path, engine='openpyxl')
                df = pd.read_excel(xls, sheet_name="美股", header=None)  # 讀取名叫 "美股" 的工作表
                
                # 取得 O10 到 P15 區間的最大值與最小值
                o10_p15 = df.iloc[9:15, 14:16]  # iloc[9:15, 14:16] 取出 O10 到 P15
                max_value = o10_p15.max().max()  # 取得最大值
                min_value = o10_p15.min().min()  # 取得最小值
                
                # 取得特定欄位的資料
                data = {
                    '代號': file_name,  # 加入 file_name
                    'ROE': df.iloc[12, 21],  # V13
                    '手調貴': max_value,
                    '手調淑': min_value,
                    '貴價': df.iloc[6, 10],    # K7
                    '淑價': df.iloc[4, 10],    # K5
                    '現價': df.iloc[2, 10],    # K3
                    '預期報酬': df.iloc[3, 10],    # K4
                    '財報': df.iloc[23, 0],   # A24
                    '檔案路徑': file_path   # 加入 file_path
                }
            except Exception as e:
                print(f"Failed to process file: {file_path}, error: {e}")
                # 如果讀取失敗，只填入 file_name，其他欄位保持空白
                data = {
                    '代號': file_name,
                    'ROE': None,
                    '手調貴': None,
                    '手調淑': None,
                    '貴價': None,
                    '淑價': None,
                    '現價': None,
                    '預期報酬': None,
                    '財報': None,
                    '檔案路徑': file_path
                }
        else:
            print(f"File not found: {file_path}")
            data = {
                '代號': file_name,
                'ROE': None,
                '手調貴': None,
                '手調淑': None,
                '貴價': None,
                '淑價': None,
                '現價': None,
                '預期報酬': None,
                '財報': None,
                '檔案路徑': file_path
            }
        all_data.append(data)
        count += 1
        
        # 每 10 筆資料寫檔一次
        if count % 10 == 0:
            combined_data = pd.DataFrame(all_data)
            if not headers_written:
                combined_data.to_excel(output_file, sheet_name='美股', index=False, engine='openpyxl')
                headers_written = True
            else:
                with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    combined_data.to_excel(writer, sheet_name='美股', index=False, header=False, startrow=writer.sheets['美股'].max_row)
            all_data = []  # 清空 all_data

    # 寫入剩餘的資料
    if all_data:
        combined_data = pd.DataFrame(all_data)
        if not headers_written:
            combined_data.to_excel(output_file, sheet_name='美股', index=False, engine='openpyxl')
        else:
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                combined_data.to_excel(writer, sheet_name='美股', index=False, header=False, startrow=writer.sheets['美股'].max_row)

def main():
    base_path = r"D:\work\me\what\company\system\資訊處理循環\tools\盈再表"  # 替換成實際的檔案路徑
    a_file_path = r"D:\work\me\what\company\system\資訊處理循環\tools\盈再表\a.xlsx"
    
    # 生成以當日日期為檔名的 xlsx 檔案
    today_date = datetime.now().strftime("%Y%m%d")
    output_file = os.path.join(base_path, f'{today_date}.xlsx')
    
    # 呼叫函數，讀取 a_file_path 檔案並將第一欄資料存入 list
    file_list = read_xls_column_to_list(a_file_path)
    
    # 處理對應的 xlsm 檔案並將資料寫入以當日日期為檔名的 xlsm 檔案
    process_files(base_path, file_list, output_file)

# 檢查是否直接執行此檔案
if __name__ == "__main__":
    main()