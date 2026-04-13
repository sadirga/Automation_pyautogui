import xlwings as xw
import time
import os
from datetime import datetime, timedelta
import configparser
from logger_setup import logger

def run_daily_sales_input(day_number, file_path):
    logger.info(f"=== input_sales started for day {day_number} ===")
    start_time = time.time()
    # Load config
    config = configparser.ConfigParser()
    config.read("config.ini")

    # Load all floor files into a dict from config
    paths = _load_config_paths(config)
    floor_files = dict(config.items("FLOORS"))
    floor_files_profit = dict(config.items("FLOORS_2"))
        
    # Open excel and load the addin (macro excel)
    app = xw.App(visible=True)
    wb_addin = app.books.open(paths["addin_path"])

    try:
        # Open daily report and get reference
        wb_daily_report = _open_workbook(file_path) #Ini path dinamis dari main.py
        ws_revenue = wb_daily_report.sheets["REVENUE"]
        ws_profit = wb_daily_report.sheets["Profit"]
        ws_input = wb_daily_report.sheets["INPUT"]
        ws_other_revenue = wb_daily_report.sheets["Other_Revenue"]
        ws_rev_sharing = wb_daily_report.sheets["RevSharing"]

        # Update day number
        ws_input["B4"].value = day_number

        # Get start column
        start_col_index = int(ws_revenue.range("F1").value) + 4

        ##### Input sheet 'REVENUE'#####
        #gf_values = _open_run_macro_get_values(app, paths["file_path_gf"], "GF", "B29:B31")
        #ws_revenue.range((6, start_col_index)).value = gf_values

        # Floors from config loop
        start_rows = {"GF":6, "1F": 11, "2F": 16, "3F": 20}
        for floor, path in floor_files.items():
            values = _open_run_macro_get_values(
                app, path, floor.upper(), "B29:B32" if floor.upper() == "1F" else "B29:B31"
            )
            ws_revenue.range((start_rows[floor.upper()], start_col_index)).value = values
        
        ##### Input Sheet 'Profit'######
        # Floors from config loop
        start_rows = {"GF_P": 6, "1F_P": 11, "2F_P": 16, "3F_P": 20}
        for floor, path in floor_files_profit.items():
            values = _open_run_macro_get_profit(
                app, path, floor.upper(), "B29:B30" if floor.upper() == "3F_P" else "B29:B31"
            )
            ws_profit.range((start_rows[floor.upper()], start_col_index)).value = values

        # Function for sheet "INPUT"
        _fill_customer_counts(app, ws_input, paths["total_sales"])    
        
        # Other revenue
        _fill_other_revenue(app, ws_input, ws_other_revenue, paths["all_brand"])
        
        # Rev_share
        _fill_rev_sharing(app, ws_rev_sharing, paths["all_brand"])
        
        wb_daily_report.save()
        wb_daily_report.close()
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
def _load_config_paths(config):
    return {
        "file_path_gf": config.get("SETTINGS", "file_path_gf"),
        "daily_report": config.get("SETTINGS", "daily_report"),
        "personal_path": config.get("SETTINGS", "personal_path"),
        "total_sales": config.get("SETTINGS", "totalSales_path"),
        "addin_path": config.get("SETTINGS", "addin_path"),
        "all_brand": config.get("SETTINGS", "allbrand"),
    } 

def _open_workbook(filepath):
    # Normalize path (important for comparisons)
    filepath = os.path.abspath(filepath)
    print("open_wb sato" + filepath)
    # Check all open workbooks in all apps
    for app in xw.apps:
        for wb in app.books:
            if os.path.abspath(wb.fullname) == filepath:
                logger.info("Workbook already open.")
                return wb

    # If not open, open it
    return xw.books.open(filepath)
        
# -----------------------------------------------------------------------------------------------------------#
#                                       Macro Helper
# -----------------------------------------------------------------------------------------------------------#
def _open_run_macro_get_values(app, wb_path, sheet_name, cell_range):
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
    
def _open_run_macro_get_profit(app, wb_path, sheet_name, cell_range):
    wb = app.books.open(wb_path)
    sheet = wb.sheets[sheet_name]
    time.sleep(1)

    app.macro("mytools.xlam!clean_daily_profit")()

    values = sheet.range(cell_range).value
    if not isinstance(values, list):
        values = [values]
    vertical_values = [[v] for v in values]
    wb.close()
    return vertical_values
    
def _open_run_macro_other_revenue(app, wb_path, sheet_name, cell_range):
    wb = app.books.open(wb_path)
    sheet = wb.sheets[sheet_name]
    time.sleep(1)

    app.macro("mytools.xlam!for_other_doc")()

    values = sheet.range(cell_range).value
    normal_values = [v for v in values]
    wb.close()
    return normal_values
    
def _fill_customer_counts(app, sheet_name, total_sales):
    # Get start column
    start_col_index = int(sheet_name.range("B4").value) + 2

    # totalSales
    wb = app.books.open(total_sales)
    ws_total_sales = wb.sheets["total_sales"]
    cust_cnt = [ws_total_sales["S10"].value,ws_total_sales["T10"].value,ws_total_sales["T11"].value]

    start_rows = {"cust_cnt": 25, "cust_tran": 29, "cust_tran_sacc": 30}
    for index, value in enumerate(start_rows):
        sheet_name.range((start_rows[value], start_col_index)).value = cust_cnt[index]
    
    wb.close()

# -----------------------------------------------------------------------------------------------------------#
#                                       Other Revenue Input
# -----------------------------------------------------------------------------------------------------------#
def _fill_other_revenue(app, ws_input, sheet_name, all_brand):
    # Get start column
    start_col_index = int(ws_input.range("B4").value) + 4
    
    # Revenue
    other_revenue_values = _open_run_macro_other_revenue(app, all_brand,"all_brand", "M2:M33") # Ini Statis, harus dibuat macro nih
    start_rows = {"440106": 4, 
    "440105" : 11, 
    "121160" : 17, 
    "511117" : 23,
    # "670106" : 203,
    # "670112" : 204, 
    # "670113" : 205,
    # "670141" : 206,
    # "670101" : 207,
    # "670102" : 208,
    # "670104" : 209,
    # "670111" : 210,    
    # "670105" : 211,
    # "670118" : 212,    
    # "670132" : 213,
    # "670140" : 214,
    # "670109" : 215,
    # "670126" : 218,
    # "670144" : 219, 
    # "670135" : 220,
    # "670138" : 221,
    # "670142" : 222,
    # "441179" : 223,
    # "441176" : 224,
    # "441193" : 225,
    # "314140" : 228,
    # "314141" : 229,
    # "551180" : 235, 
    # "124150" : 239,
    # "124136" : 240,
    # "124137" : 242,
    # "124139" : 243, 
    }
    for i, value in enumerate(start_rows):
        sheet_name.range((start_rows[value], start_col_index)).value = other_revenue_values[i]
   
def _fill_rev_sharing(app, sheet_name, all_brand):
    # Get start column
    start_col_index = 15
    
    # Revenue
    Rev_val = _open_run_macro_other_revenue(app, all_brand,"all_brand", "M6:M47") # Ini Statis, harus dibuat macro nih
    start_rows = {
    "670106" : 6,
    "670112" : 7, 
    "670113" : 8,
    "670141" : 9,
    "670101" : 10,
    "670102" : 11,
    "670104" : 12,
    "670111" : 13,    
    "670105" : 14,
    "670118" : 15,    
    "670132" : 16,
    "670140" : 17,
    "670109" : 18,
    "670126" : 19,
    "670144" : 20, 
    "670135" : 21,
    "670138" : 22,
    "670142" : 23,
    "441179" : 24,
    "441176" : 25,
    "441193" : 26,
    "314140" : 27,
    "314141" : 28,
    "551180" : 29, 
    "124150" : 30,
    "124136" : 31,
    "124137" : 32,
    "124139" : 33, 
    "441177" : 34,
    "670121" : 35,
    "155106" : 36,
    "542108" : 37,
    "542112" : 38,
    "553105" : 39,
    "553111" : 40,
    "553112" : 41,
    "553113" : 42,
    "553114" : 43,
    "553115" : 44,
    "553116" : 45,
    "553117" : 46,
    "553118" : 47, 
    
    }
    for i, value in enumerate(start_rows):
        sheet_name.range((start_rows[value], start_col_index)).value = Rev_val[i]    