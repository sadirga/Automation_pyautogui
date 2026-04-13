import xlwings as xw
import time
from datetime import datetime, timedelta
import configparser
import calendar
import os
from tkcalendar import Calendar
import tkinter as tk
from logger_setup import logger

def get_or_open_wb(filepath):
    filepath = os.path.abspath(filepath)

    # Try to find the file among open Excel apps
    for app in xw.apps:
        for wb in app.books:
            if os.path.abspath(wb.fullname) == filepath:
                logger.info(f"Workbook already open → {filepath}")
                return wb, app, False   # (workbook, app, newly opened? False)

    # Not open → open fresh Excel instance
    logger.info(f"Opening workbook → {filepath}")
    app = xw.App(visible=True)
    wb = app.books.open(filepath)
    return wb, app, True   # newly opened

def get_file_path(selected_date: datetime) -> str: # Agar dinamis nama file nya 
    base_folder = config["SETTINGS"]["base_folder"]
    month_num   = selected_date.month              # e.g. 3
    month_abbr  = selected_date.strftime("%b")     # e.g. Mar
    year        = selected_date.year               # e.g. 2026

    filename = f"{month_num}. Sales Report {month_abbr} {year}.xlsm"
    print(filename)
    return os.path.join(base_folder, filename)   

def ask_for_date():
    selected = []

    root = tk.Tk()
    root.title("Pilih Tanggal")
    root.resizable(False, False)

    today = datetime.today() - timedelta(days=1) if datetime.today().day > 1 else datetime.today()

    cal = Calendar(
        root,
        selectmode="day",
        year=today.year,
        month=today.month,
        day=today.day,
        date_pattern="yyyy-mm-dd",
    )
    cal.pack(padx=20, pady=20)

    def confirm():
        selected.append(cal.get_date())
        root.destroy()

    tk.Button(root, text="Konfirmasi", command=confirm).pack(pady=(0, 15))
    root.grab_set()
    root.lift()              # bring window to front
    root.focus_force()       # force keyboard focus
    root.attributes("-topmost", True)   # stay on top of all windows
    root.mainloop()

    if not selected:
        raise ValueError("Tanggal tidak dipilih.")

    return datetime.strptime(selected[0], "%Y-%m-%d")
        
def daily_2(wb_path, sheet_name, sheet_name_2):   
    
    config = configparser.ConfigParser()
    config.read("config.ini")

    file_path = config.get("SETTINGS", "daily_report_2_path")
    base_filename = config.get("SHEETS", "daily_report_2_name")
    dr2_sheet_name = config.get("SHEETS", "dr2_sheet_name")

    # Build filename safely
    input_date = ask_for_date()
    yesterday = input_date.strftime("%y%m%d")
    #yesterday = (datetime.now() - timedelta(days=1)).strftime("%y%m%d")
    year_input = input_date.strftime("%y")
    month_input = input_date.month
    last_day = calendar.monthrange(datetime.now().year, datetime.now().month)[1]
    
    front_name = f"{yesterday}★ {year_input}년 {month_input}월 "
    final_filename = os.path.join(file_path, f"{front_name}{base_filename}")
    
    # Open DR2 workbook
    wb, app, new_app = get_or_open_wb(final_filename)
    ws = wb.sheets[dr2_sheet_name]

    # Open daily workbook
    daily_wb, daily_app, new_daily_app = get_or_open_wb(wb_path)
    date_input = input_date.day
    #date_input = daily_wb.sheets[sheet_name].range("B4").value
    print(last_day,yesterday, date_input)
    col = int(date_input) + 4
    row = 35

    input_value = daily_wb.sheets[sheet_name_2].range((row, col)).value

    # Last row logic
    last_row = ws.range("BC" + str(ws.cells.last_cell.row)).end("up").row
    input_row = int(last_row - (last_day - date_input)) - 1

    # Targets
    target_cell = ws.range((last_row - last_day - 1, 54))
    target_cell2 = ws.range((last_row - last_day - 1, 57))

    # Formulas
    formula = f"=SUM(BB{last_row - last_day}:BB{input_row})"
    formula2 = f"=SUM(BE{last_row - last_day}:BE{input_row})"

    target_cell.formula = formula
    target_cell2.formula = formula2

    ws[f"BC{input_row}"].value = input_value

    logger.info(f"input row: {input_row}, value: {input_value}, last_row: {last_row}")

    # Save workbook
    wb.save()

    # Only close app if we created it
    if new_app:
        wb.close()
        app.quit()

    if new_daily_app:
        daily_wb.close()
        daily_app.quit()

def main():
    # Start time
    start = time.time()
    logger.info("Script started")
    
    # Load config
    config = configparser.ConfigParser()
    config.read("config.ini")   
    
    daily_sales_report = config.get("SETTINGS", "daily_report")
    input_sheet = config.get("SHEETS", "input_sheet")
    revenue_sheet = config.get("SHEETS", "revenue_sheet")
    
    # Running script
    daily_2(daily_sales_report, input_sheet, revenue_sheet)
    
    # End time
    elapsed = time.time() - start
    logger.info(f"Script running succesfully, finished in {elapsed:.2f} seconds")
    
if __name__ == "__main__":
    main()