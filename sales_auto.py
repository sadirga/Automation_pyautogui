def input_sales():
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

    # Load all floor files into a dict from config
    floor_files = dict(config.items("FLOORS"))
    
    # Using already open personal.xlb
    app = xw.apps.active

    # Run macro helper
    def run_macro_and_get_values(wb_path, sheet_name, cell_range):
        wb = app.books.open(wb_path)
        sheet = wb.sheets[sheet_name]
        time.sleep(1)

        macro = app.books["PERSONAL.XLSB"].macro("clean_daily")
        macro()

        values = sheet.range(cell_range).value
        if not isinstance(values, list):
            values = [values]
        vertical_values = [[v] for v in values]
        wb.close()
        return vertical_values

    # 1️⃣ Open daily report and get reference
    daily = app.books.open(daily_report)
    daily_sheet = daily.sheets["REVENUE"]
    input_sheet = daily.sheets["INPUT"]

    # Update day number
    day_number = (datetime.now() - timedelta(days=1)).day
    input_sheet["B4"].value = day_number

    # Get start column
    start_col_index = int(daily_sheet.range("F1").value) + 4

    # 2️⃣ GF
    gf_values = run_macro_and_get_values(file_path_gf, "GF", "B29:B31")
    daily_sheet.range((6, start_col_index)).value = gf_values

    # 3️⃣ Floors from config loop
    start_rows = {"1f": 11, "2f": 17, "3f": 21}
    for floor, path in floor_files.items():
        values = run_macro_and_get_values(path, floor.upper(), "B29:B33" if floor == "1f" else "B29:B31")
        daily_sheet.range((start_rows[floor], start_col_index)).value = values