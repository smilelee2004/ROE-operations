# 將 MonitorList.xlsx 檔案裡面的 stock ID , 依次把每個 Stock ID 的盈再表裡的關鍵數值, 輸出到yyyymmdd.xlsx 裡面, 方便複製到 monitorlist 裡面.async 
import pandas as pd
import os
from datetime import datetime
from openpyxl import load_workbook
from tqdm import tqdm
import threading
import tkinter as tk
from tkinter import messagebox
from tkinter import ttk  # 導入 ttk 模組
from tkinter import filedialog  # 導入資料夾選擇對話框

def read_xls_column_to_list(file_path):
    # 讀取 xlsm 檔案
    wb = load_workbook(file_path, read_only=True)
    sheet = wb.active

    # 取得第一欄 (Column 1) 的資料並轉換成 list
    column_1_list = [row[0] for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)]

    return column_1_list

def find_monitorlist(base_path):
    """在指定目錄中尋找 monitorlist 檔案 (不分大小寫，支援 .xlsx / .xlsm / .xls)。
    找到回傳完整路徑，找不到回傳 None。"""
    if not os.path.isdir(base_path):
        return None
    # 依優先順序檢查副檔名
    preferred_exts = ['.xlsx', '.xlsm', '.xls']
    candidates = []
    for entry in os.listdir(base_path):
        name, ext = os.path.splitext(entry)
        # 檔名以 monitorlist 開頭 (不分大小寫)，且副檔名為 Excel 格式
        if name.lower().startswith('monitorlist') and ext.lower() in preferred_exts:
            candidates.append(entry)
    if not candidates:
        return None
    # 依副檔名優先順序排序後取第一個
    candidates.sort(key=lambda f: preferred_exts.index(os.path.splitext(f)[1].lower()))
    return os.path.join(base_path, candidates[0])

# 盈再表檔名可能帶有來源後綴，例如 AAPL_SEC.xlsx 或 ABBNY_Yahoo.xlsx。
# 依優先順序列出要嘗試的後綴；'' 代表舊格式 (無後綴，例如 AAPL.xlsx)。
ROE_SOURCE_SUFFIXES = ['_SEC', '_Yahoo', '']
ROE_FILE_EXTS = ['.xlsx', '.xlsm']

def find_roe_file(base_path, file_name):
    """在指定目錄中尋找某個代號的盈再表檔案。

    檔名可能帶來源後綴 (例如 {代號}_SEC.xlsx、{代號}_Yahoo.xlsx)，
    也相容舊格式的 {代號}.xlsx / {代號}.xlsm。

    比對方式：
      1. 依 ROE_SOURCE_SUFFIXES x ROE_FILE_EXTS 的優先順序做「精確檔名」比對，
         確保代號是完整比對 (例如 'AN' 不會誤抓到 'ANZGY_Yahoo.xlsx')。
      2. 若上述都找不到，退而以 {代號}_*.xlsx / {代號}_*.xlsm 比對其他來源後綴。
    找到回傳完整路徑；都找不到時回傳預設的 _SEC.xlsx 路徑 (供錯誤訊息使用)。"""
    import glob

    code = str(file_name)

    # 1. 已知後綴的精確比對 (依優先順序)
    for suffix in ROE_SOURCE_SUFFIXES:
        for ext in ROE_FILE_EXTS:
            candidate = os.path.join(base_path, f"{code}{suffix}{ext}")
            if os.path.exists(candidate):
                return candidate

    # 2. 後備：比對任何 {代號}_其他來源 的檔案 (底線確保不會誤判前綴)
    for ext in ROE_FILE_EXTS:
        pattern = os.path.join(glob.escape(base_path), f"{glob.escape(code)}_*{ext}")
        matches = sorted(glob.glob(pattern))
        if matches:
            return matches[0]

    # 都找不到，回傳預設路徑 (os.path.exists 仍會是 False)
    return os.path.join(base_path, f"{code}_SEC.xlsx")

def process_files(base_path, file_list, output_file, progress_var, cancel_event, root):
    all_data = []
    abnormal_data = []
    abnormalFlag = False
    count = 0
    headers_written = False

    for file_name in tqdm(file_list, desc="Processing files", unit="file"):
        if cancel_event.is_set():
            print("Processing cancelled")
            break

        # 依代號尋找盈再表檔案，優先 .xlsm，找不到再找 .xlsx
        file_path = find_roe_file(base_path, file_name)

        if os.path.exists(file_path):
#            #print(f"Processing file: {file_path}")
            try:
                # 讀取對應的 xlsm 檔案
                wb = load_workbook(file_path, data_only=True, read_only=False)
                sheet = wb["美股"]  # 讀取名叫 "美股" 的工作表
                
                # 取得 O10 到 P15 區間的所有數值，排除 None
                values = [cell.value for row in sheet.iter_rows(min_row=10, max_row=15, min_col=15, max_col=16) for cell in row if cell.value is not None]
                if values:
                    max_value = max(values)
                    min_value = min(values)
                    abnormalFlag = False
                else:
                    max_value = None
                    min_value = None
                    abnormalFlag = True
                
#                #print(f"ROE {sheet.cell(row=13, column=22).value}")
                # 取得特定欄位的資料
                data = {
                    '代號': file_name,  # 加入 file_name
                    'ROE': sheet.cell(row=13, column=22).value,  # V13
                    '手調貴': max_value,
                    '手調淑': min_value,
                    '貴價': sheet.cell(row=7, column=11).value,    # K7
                    '淑價': sheet.cell(row=5, column=11).value,    # K5
                    '現價': sheet.cell(row=3, column=11).value,    # K3
                    '預期報酬': sheet.cell(row=4, column=11).value,    # K4
                    '財報': sheet.cell(row=24, column=1).value,   # A24
                    '檔案路徑': f'=HYPERLINK("{file_path}", "點我開啟檔案")'    # 加入 file_path
                }
                
            except Exception as e:
                print(f"Failed to process file: {file_path}, error: {e}")
                # 如果讀取失敗，只填入 file_name，其他欄位保持空白
                data = {
                    '代號': file_name,
                    'ROE': sheet.cell(row=13, column=22).value,  # V13,
                    '手調貴': None,
                    '手調淑': None,
                    '貴價': sheet.cell(row=7, column=11).value,    # K7
                    '淑價': sheet.cell(row=5, column=11).value,    # K5
                    '現價': sheet.cell(row=3, column=11).value,    # K3
                    '預期報酬': sheet.cell(row=4, column=11).value,    # K4
                    '財報': sheet.cell(row=24, column=1).value,   # A24
                    '檔案路徑': f'=HYPERLINK("{file_path}", "點我開啟檔案")'
                }
                abnormalFlag = True

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
                '檔案路徑': f'=HYPERLINK("{file_path}", "點我開啟檔案")'
            }
            abnormalFlag = True
        
        all_data.append(data)
        if abnormalFlag or data.get('預期報酬') == 'na':
            # 如果有異常，或預期報酬為 'na'，將資料加入 abnormal_data
            abnormal_data.append(data)
        
        count += 1
        
        # 更新進度條
        progress_var.set(count)
        root.update_idletasks()
        
        # 每 10 筆資料寫檔一次
        if count % 10 == 0:
            combined_data = pd.DataFrame(all_data)
            combined_abnormal = pd.DataFrame(abnormal_data)
            if not headers_written:
                combined_data.to_excel(output_file, sheet_name='美股', index=False, engine='openpyxl')
                headers_written = True
            else:
                with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                    combined_data.to_excel(writer, sheet_name='美股', index=False, header=False, startrow=writer.sheets['美股'].max_row)
            all_data = []  # 清空 all_data

             # 寫入異常資料
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                if '異常資料' in writer.sheets:
                    startrow = writer.sheets['異常資料'].max_row
                    combined_abnormal.to_excel(writer, sheet_name='異常資料', index=False, header=False, startrow=startrow)
                else:
                    combined_abnormal.to_excel(writer, sheet_name='異常資料', index=False)
                #combined_abnormal.to_excel(writer, sheet_name='異常資料', index=False, header=False, startrow=writer.sheets['異常資料'].max_row)
            abnormal_data = []  # 清空 abnormal_data

    # 寫入剩餘的資料
    if all_data:
        combined_data = pd.DataFrame(all_data)
        if not headers_written:
            combined_data.to_excel(output_file, sheet_name='美股', index=False, engine='openpyxl')
        else:
            with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                combined_data.to_excel(writer, sheet_name='美股', index=False, header=False, startrow=writer.sheets['美股'].max_row)
    if abnormal_data:
        combined_abnormal = pd.DataFrame(abnormal_data)
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
            if '異常資料' in writer.sheets:
                startrow = writer.sheets['異常資料'].max_row
                combined_abnormal.to_excel(writer, sheet_name='異常資料', index=False, header=False, startrow=startrow)
            else:
                combined_abnormal.to_excel(writer, sheet_name='異常資料', index=False)

def main():
    # 預設目錄 (可保留作為對話框的起始位置)
    default_path = r"D:\work\me\what\company\system\資訊處理循環\tools\盈再表"
    initial_dir = default_path if os.path.isdir(default_path) else os.getcwd()

    # 彈出資料夾選擇視窗，讓使用者指定盈再表來源目錄
    picker_root = tk.Tk()
    picker_root.withdraw()  # 隱藏多餘的主視窗，只顯示對話框
    base_path = filedialog.askdirectory(
        title="請選擇盈再表來源目錄 (內含 MonitorList 與各代號盈再表)",
        initialdir=initial_dir
    )
    picker_root.destroy()

    # 使用者取消選擇
    if not base_path:
        print("未選擇目錄，程式結束。")
        return

    base_path = os.path.normpath(base_path)

    # 在選定目錄中自動偵測 monitorlist 檔案
    a_file_path = find_monitorlist(base_path)
    if not a_file_path:
        # 沒有 GUI 主視窗也能彈出訊息
        err_root = tk.Tk()
        err_root.withdraw()
        messagebox.showerror(
            "找不到 MonitorList",
            f"在目錄中找不到 MonitorList 檔案 (.xlsx / .xlsm / .xls)：\n{base_path}"
        )
        err_root.destroy()
        print(f"找不到 MonitorList 檔案於：{base_path}")
        return

    print(f"來源目錄：{base_path}")
    print(f"MonitorList 檔案：{a_file_path}")

    # 生成以當日日期為檔名的 xlsx 檔案
    today_date = datetime.now().strftime("%Y%m%d")
    output_file = os.path.join(base_path, f'{today_date}.xlsx')

    # 呼叫函數，讀取 a_file_path 檔案並將第一欄資料存入 list
    file_list = read_xls_column_to_list(a_file_path)

    # 創建進度條和取消按鈕
    root = tk.Tk()
    root.title("盈再表 --> Monoitor list Processing ...")
    
    progress_var = tk.DoubleVar()
    progress_bar = ttk.Progressbar(root, variable=progress_var, maximum=len(file_list))  # 使用 ttk.Progressbar
    progress_bar.pack(fill=tk.X, expand=1, padx=10, pady=10)
    
    cancel_event = threading.Event()
    
    def cancel():
        cancel_event.set()
        messagebox.showinfo("Cancelled", "Processing has been cancelled.")
    
    cancel_button = tk.Button(root, text="Cancel", command=cancel)
    cancel_button.pack(pady=10)
    
    def run_processing():
        process_files(base_path, file_list, output_file, progress_var, cancel_event, root)
        root.quit()
    
    threading.Thread(target=run_processing).start()
    root.mainloop()

# 檢查是否直接執行此檔案
if __name__ == "__main__":
    main()