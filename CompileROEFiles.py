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
    """讀 MonitorList 第一欄的代號清單，只取真正有值的列。

    為什麼要濾：Excel 的「已使用範圍」常被撐得比實際資料長很多（儲存格編輯過
    再清空就會留下痕跡），iter_rows 會一路讀到那裡，尾巴全是 None。
    盈再表_輸出\\MonitorList.xlsx 就是讀出 720 列、其中 281 列是 None，
    而 find_roe_file() 會把 None 用 str() 變成字串 "None" 去組檔名，
    於是刷出 281 行 "File not found: ...\\None_SEC.xlsx"。
    """
    wb = load_workbook(file_path, read_only=True)
    try:
        sheet = wb.active
        codes, seen = [], set()
        for row in sheet.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True):
            value = row[0]
            if value is None:
                continue
            code = str(value).strip()
            if not code or code.lower() in ("none", "nan"):
                continue
            if code in seen:          # 清單偶有重複，重複跑只是白做工
                continue
            seen.add(code)
            codes.append(code)
        return codes
    finally:
        wb.close()

# MonitorList 的唯一權威位置。
# 2026-09-05 之前，三支工具各讀各的副本：run_full_pipeline 讀 盈再表\、
# 本程式讀使用者選的目錄（實務上是 盈再表_輸出\）、tools\ 那份反而沒人讀，
# 而且三份內容不一樣（多 CCL / 少 CUK、HOLX，多已下市的 RBGLD …），
# 導致「pipeline 跑的標的」與「收尾總表的標的」長期對不起來。
# 現在一律以 tools\MonitorList.xlsx 為準；可用環境變數 YZB_MONITORLIST 覆寫。
WORKSPACE_MONITORLIST = os.environ.get(
    "YZB_MONITORLIST",
    r"D:\work\me\what\company\system\資訊處理循環\tools\MonitorList.xlsx")


def find_monitorlist(base_path):
    """取得 MonitorList 路徑。

    優先用唯一權威位置 (WORKSPACE_MONITORLIST)；那份不存在時，才退回舊行為
    ——在使用者選的目錄裡找 monitorlist*.xlsx/.xlsm/.xls。
    找不到回傳 None。"""
    if os.path.isfile(WORKSPACE_MONITORLIST):
        return WORKSPACE_MONITORLIST
    print(f"[警告] 找不到權威 MonitorList：{WORKSPACE_MONITORLIST}")
    print(f"       退回在所選目錄尋找：{base_path}")
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

def blank_row(code, file_path):
    """讀不到內容時的空白列 —— 仍保留代號，讓輸出看得出這一檔沒抓到。"""
    return {
        '代號': code,
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
            wb = None
            try:
                # 讀取對應的 xlsm 檔案
                wb = load_workbook(file_path, data_only=True, read_only=False)
                if "美股" not in wb.sheetnames:
                    raise KeyError(f"活頁簿沒有『美股』工作表 (實際有: {wb.sheetnames})")
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
                # 舊版這裡照樣去讀 sheet.cell(...)，但若爆的是 load_workbook 或
                # wb["美股"]，sheet 根本還沒被賦值 → except 內再拋 NameError，
                # 反而把真正的錯誤訊息蓋掉。改成完全不依賴 sheet。
                print(f"Failed to process file: {file_path}, error: {type(e).__name__}: {e}")
                data = blank_row(file_name, file_path)
                abnormalFlag = True
            finally:
                # 舊版從不關檔；440 個 790KB 的活頁簿累積下來，檔案握把與記憶體都會漲。
                if wb is not None:
                    try:
                        wb.close()
                    except Exception:
                        pass

        else:
            print(f"File not found: {file_path}")
            data = blank_row(file_name, file_path)
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

    print(f"盈再表來源目錄：{base_path}")
    print(f"MonitorList    ：{a_file_path}")
    if os.path.normcase(os.path.abspath(a_file_path)) != os.path.normcase(os.path.abspath(WORKSPACE_MONITORLIST)):
        print("  ⚠ 這不是權威位置的 MonitorList，標的清單可能與 pipeline 跑的不一致")

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