import xlwings as xw
import time
import os
from datetime import datetime, timedelta
import configparser
from logger_setup import logger

def input_sales(day_number):
    logger.info(f"=== input_sales started for day {day_number} ===")
    start_time = time.time()
    # Load config
    config = configparser.ConfigParser()
    config.read("config.ini")

    # Load all floor files into a dict from config
    paths = _load_paths(config)
    floor_files = dict(config.items("FLOORS"))
        
    # Open excel and load the addin (macro excel)
    app = xw.App(visible=True)
    addin = app.books.open(paths["addin_path"])

    try:
        # Open daily report and get reference
        daily = _get_or_open_wb(paths["daily_report"])
        daily_sheet = daily.sheets["REVENUE"]
        input_sheet = daily.sheets["INPUT"]
        other_sheet = daily.sheets["Other_Revenue"]

        # Update day number
        input_sheet["B4"].value = day_number

        # Get start column
        start_col_index = int(daily_sheet.range("F1").value) + 4

        # GF
        gf_values = _run_macro_and_get_values(app, paths["file_path_gf"], "GF", "B29:B31")
        daily_sheet.range((6, start_col_index)).value = gf_values

        # Floors from config loop
        start_rows = {"1f": 11, "2f": 17, "3f": 21}
        for floor, path in floor_files.items():
            values = _run_macro_and_get_values(
                app, path, floor.upper(), "B29:B33" if floor == "1f" else "B29:B31"
            )
            daily_sheet.range((start_rows[floor], start_col_index)).value = values
        
        # Function for sheet "INPUT"
        _input_sheet_sales(app, input_sheet, paths["total_sales"])    
        
        # Other revenue
        _other_revenue_sales(app, input_sheet, other_sheet, paths["all_brand"])
        
        daily.save()
        daily.close()
        logger.info("Daily report saved and closed")
        
    except Exception:
        logger.exception("Error during input_sales execution")
        raise
        
    finally:
        app.quit()
        elapsed = time.time() - start_time
        mins, secs = divmod(int(elapsed), 60)
        logger.info(f"=== input_sales finished in {mins}m {secs}s ===")
        
# -----------------------------------------------------------------------------------------------------------#
#                                       Private helper functions
# -----------------------------------------------------------------------------------------------------------#
def _load_paths(config):
    return {
        "file_path_gf": config.get("SETTINGS", "file_path_gf"),
        "daily_report": config.get("SETTINGS", "daily_report"),
        "personal_path": config.get("SETTINGS", "personal_path"),
        "total_sales": config.get("SETTINGS", "totalSales_path"),
        "addin_path": config.get("SETTINGS", "addin_path"),
        "all_brand": config.get("SETTINGS", "allbrand"),
    } 

def _get_or_open_wb(filepath):
    # Normalize path (important for comparisons)
    filepath = os.path.abspath(filepath)

    # Check all open workbooks in all apps
    for app in xw.apps:
        for wb in app.books:
            if os.path.abspath(wb.fullname) == filepath:
                logger.info("Workbook already open.")
                return wb

    # If not open, open it
    return xw.books.open(filepath)
        
# Run macro helper
def _run_macro_and_get_values(app, wb_path, sheet_name, cell_range):
    wb = app.books.open(wb_path)
    sheet = wb.sheets[sheet_name]
    time.sleep(1)

    app.macro("mytools.xlam!clean_daily")()

    values = sheet.range(cell_range).value
    if not isinstance(values, list):
        values = [values]
    vertical_values = [[v] for v in values]
    wb.close()
    return vertical_values
    
def _get_values_other(app, wb_path, sheet_name, cell_range):
    wb = app.books.open(wb_path)
    sheet = wb.sheets[sheet_name]
    time.sleep(1)

    app.macro("mytools.xlam!for_other_doc")()

    values = sheet.range(cell_range).value
    normal_values = [v for v in values]
    wb.close()
    return normal_values
    
def _input_sheet_sales(app, sheet_name, total_sales):
    # Get start column
    start_col_index = int(sheet_name.range("B4").value) + 2

    # totalSales
    wb = app.books.open(total_sales)
    ts_sheet = wb.sheets["total_sales"]
    cust_cnt = [ts_sheet["S10"].value,ts_sheet["T10"].value,ts_sheet["T11"].value]

    start_rows = {"cust_cnt": 25, "cust_tran": 29, "cust_tran_sacc": 30}
    for index, value in enumerate(start_rows):
        sheet_name.range((start_rows[value], start_col_index)).value = cust_cnt[index]
    
    wb.close()

def _other_revenue_sales(app, input_sheet, sheet_name, all_brand):
    # Get start column
    start_col_index = int(input_sheet.range("B4").value) + 4
    
    # Revenue
    other_values = _get_values_other(app, all_brand,"all_brand", "M2:M5")
    start_rows = {"440106": 4, "440105" : 11, "121160" : 17, "511117" : 23}
    for i, value in enumerate(start_rows):
        sheet_name.range((start_rows[value], start_col_index)).value = other_values[i]
    