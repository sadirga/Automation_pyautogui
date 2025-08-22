def input_sales(day_number):
    import xlwings as xw
    import time
    from datetime import datetime, timedelta
    import configparser

    # Load config
    config = configparser.ConfigParser()
    config.read("config.ini")

    file_path_gf = config.get("SETTINGS", "file_path_gf")
    daily_report = config.get("SETTINGS", "daily_report")
    personal_path = config.get("SETTINGS", "personal_path")
    totalSales = config.get("SETTINGS", "totalSales_path")
    addin_path = config.get("SETTINGS", "addin_path")
    all_brand = config.get("SETTINGS", "allbrand")

    # Load all floor files into a dict from config
    floor_files = dict(config.items("FLOORS"))
    
    ### Using already open personal.xlb ###
    ### app = xw.apps.active ### previous code to activate excel from existing open workbook
    
    # Open excel and load the addin (macro excel)
    app = xw.App(visible=True)
    addin = app.books.open(addin_path)
    
    # Run macro helper
    def run_macro_and_get_values(wb_path, sheet_name, cell_range):
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
        
    def get_values_other(wb_path, sheet_name, cell_range):
        wb = app.books.open(wb_path)
        sheet = wb.sheets[sheet_name]
        time.sleep(1)

        app.macro("mytools.xlam!for_other_doc")()

        values = sheet.range(cell_range).value
        normal_values = [v for v in values]
        wb.close()
        return normal_values
        
    def input_sheet_sales(sheet_name):
        # Get start column
        start_col_index = int(sheet_name.range("B4").value) + 2

        # totalSales
        wb = app.books.open(totalSales)
        ts_sheet = wb.sheets["totalSales"]
        cust_cnt = [ts_sheet["S10"].value,ts_sheet["T10"].value,ts_sheet["T11"].value]

        start_rows = {"cust_cnt": 25, "cust_tran": 29, "cust_tran_sacc": 30}
        for index, value in enumerate(start_rows):
            sheet_name.range((start_rows[value], start_col_index)).value = cust_cnt[index]
        
        wb.close()
    
    def other_revenue_sales(input_sheet, sheet_name):
        # Get start column
        start_col_index = int(input_sheet.range("B4").value) + 4
        
        # Revenue
        other_values = get_values_other(all_brand,"allbrand", "M2:M5")
        start_rows = {"440106": 4, "440105" : 11, "121160" : 17, "511117" : 23}
        for i, value in enumerate(start_rows):
            sheet_name.range((start_rows[value], start_col_index)).value = other_values[i]
        

    # Open daily report and get reference
    daily = app.books.open(daily_report)
    daily_sheet = daily.sheets["REVENUE"]
    input_sheet = daily.sheets["INPUT"]
    other_sheet = daily.sheets["Other_Revenue"]

    # Update day number
    input_sheet["B4"].value = day_number

    # Get start column
    start_col_index = int(daily_sheet.range("F1").value) + 4

    # GF
    gf_values = run_macro_and_get_values(file_path_gf, "GF", "B29:B31")
    daily_sheet.range((6, start_col_index)).value = gf_values

    # Floors from config loop
    start_rows = {"1f": 11, "2f": 17, "3f": 21}
    for floor, path in floor_files.items():
        values = run_macro_and_get_values(path, floor.upper(), "B29:B33" if floor == "1f" else "B29:B31")
        daily_sheet.range((start_rows[floor], start_col_index)).value = values
    
    # Function for sheet "INPUT"
    input_sheet_sales(input_sheet)    
    
    # Other revenue
    other_revenue_sales(input_sheet, other_sheet)
    
    daily.save()
    daily.close()
    app.quit()
    