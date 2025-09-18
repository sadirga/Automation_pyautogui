import pyautogui
import psutil
import time
import sales_auto as sato
from datetime import datetime
import ctypes
import os
import tkinter as tk
from tkinter import simpledialog
import configparser
from logger_setup import logger

# -----------------------------------------------------------------------------------------------------------#
#                                       Config Helpers
# -----------------------------------------------------------------------------------------------------------#
config = configparser.ConfigParser()
config.read("config.ini")

def get_tuple(section, key):
    """Return a tuple of ints from config, e.g. '1256,622' -> (1256, 622)"""
    return tuple(int(v) for v in config.get(section, key).split(","))

def get_rgb(section, key):
    """Return (x, y, r, g, b) from config"""
    values = [int(v) for v in config.get(section, key).split(",")]
    return (values[0], values[1]), tuple(values[2:])

# -----------------------------------------------------------------------------------------------------------#
#                                       Input Helpers
# -----------------------------------------------------------------------------------------------------------#

def ask_for_date():
    root = tk.Tk()
    root.withdraw()  # hide main window
    day_str = simpledialog.askstring("Input", "Enter day number (1-31):")
    root.destroy()
    today = datetime.today()
    return datetime(today.year, today.month, int(day_str))

def get_password():
    password = os.getenv("RIS_PASSWORD")
    if not password:
        root = tk.Tk()
        root.withdraw()  # Hide the main tkinter window
        password = simpledialog.askstring("Password Required", "Enter RIS password:", show='*')
        root.destroy()
        if not password:
            raise ValueError("Password was not provided!")
    return password

# -----------------------------------------------------------------------------------------------------------#
#                                       Wait Helpers
# -----------------------------------------------------------------------------------------------------------#
    
def wait_for_pixel(x, y, expected_rgb, timeout=30):
    """Wait until pixel at (x,y) matches expected_rgb, or timeout"""
    start = time.time()
    while time.time() - start < timeout:
        pixel = pyautogui.screenshot().getpixel((x, y))
        if pixel == expected_rgb:
            return True
        time.sleep(0.5)
    raise TimeoutError(f"UI not ready after {timeout}s at pixel {x},{y}")

def wait_for_process(proc_name, timeout=30):
    """Wait until a process with given name appears."""
    start = time.time()
    while time.time() - start < timeout:
        for p in psutil.process_iter(["name"]):
            if p.info["name"] and proc_name.lower() in p.info["name"].lower():
                return p
        time.sleep(0.5)
    raise TimeoutError(f"Process '{proc_name}' not found after {timeout}s")

# -----------------------------------------------------------------------------------------------------------#
#                                       Core Automation
# -----------------------------------------------------------------------------------------------------------#

def run_app_and_login(shortcut_path, proc_name):
    # Start via shortcut
    os.startfile(shortcut_path)

    # Wait for process showing
    proc = wait_for_process(proc_name, timeout=20)

    # Wait for UI to show
    (x, y), rgb = get_rgb("PIXELS", "login_pixel")
    wait_for_pixel(x, y, rgb) # adjustable in config to match other device

    # Enter password
    password = get_password()
    pyautogui.typewrite(password, interval=0.05)
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(0.5)

def _open_sales_analysis():
    for key in ["open_sales_analysis_1", "open_sales_analysis_2", "open_sales_analysis_3"]:
        pyautogui.moveTo(*get_tuple("COORDS", key), duration=0.2)
        time.sleep(0.3)
        pyautogui.click()
        pyautogui.click()
        time.sleep(1)

def _daily_sales_analysis(input_day):
    for key in ["daily_sales_1", "daily_sales_2"]:
        pyautogui.moveTo(*get_tuple("COORDS", key), duration=0.2)
        time.sleep(0.5)
        pyautogui.click()
        pyautogui.click()
        time.sleep(1)
    
    # Set Date
    pyautogui.moveTo(*get_tuple("COORDS", "date_field"), duration=0.2)
    pyautogui.click()
    time.sleep(0.5)
    pyautogui.typewrite(input_day, interval=0.08)
    pyautogui.typewrite(input_day, interval=0.08)
    
    # Open feature
    pyautogui.moveTo(*get_tuple("COORDS", "pay_summary_tab"), duration=0.2) # Cursor memilih tab "Pay Sumamry"
    pyautogui.click()
    time.sleep(0.5)
    pyautogui.press('f2')
    time.sleep(0.5)
    
    # Wait until load daily data finished
    (x, y), rgb = get_rgb("PIXELS", "loading_pixel")
    wait_for_pixel(x, y, rgb)
    
    # Save File
    pyautogui.press('f8')
    pyautogui.typewrite("all_brand", interval=0.05)
    pyautogui.press('enter')
    # Notice after each command I slip a sleep time just for extra caution
    time.sleep(0.5)
    pyautogui.press('left')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)

def _set_date_and_options(input_day_str):
    # Setting input_day_str date as the targeted date
    pyautogui.typewrite(input_day_str, interval=0.08)
    time.sleep(1)

    # Option Button
    pyautogui.moveTo(*get_tuple("COORDS", "option_button"), duration=0.3)
    time.sleep(0.5)
    pyautogui.click()

    # Unit Rp
    pyautogui.moveTo(*get_tuple("COORDS", "unit_button"), duration=0.3)
    time.sleep(0.5)
    pyautogui.click()

    # OK Button
    pyautogui.moveTo(*get_tuple("COORDS", "ok_button"), duration=0.3)
    time.sleep(0.5)
    pyautogui.click()

def _save_sales_report(floor_name, coord_key):

    # Floor Selection
    pyautogui.moveTo(*get_tuple("COORDS", coord_key), duration=0.3)
    time.sleep(0.3)
    pyautogui.click()
    pyautogui.click()
    time.sleep(0.3)

    # Save File
    pyautogui.press('F8')
    time.sleep(0.5)
    pyautogui.typewrite(floor_name, interval=0.05)
    pyautogui.press('enter')
    time.sleep(0.3)
    pyautogui.press('left')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.5)

    # Go back to home
    pyautogui.moveTo(*get_tuple("COORDS", "back_button"), duration=0.5)
    time.sleep(0.3)
    pyautogui.click()
    pyautogui.click()
        
def run_sales_automation():
    start_time = time.time()
    logger.info("=== Sales automation started ===")
    
    try:    
        # Go to desktop Windows + D
        pyautogui.hotkey("winleft", "d")
        time.sleep(2)
        
        app_path = config.get("SETTINGS", "app_path")
        app_name = config.get("SETTINGS", "app_name")
       
        input_date = ask_for_date()
        input_day_str = input_date.strftime("%Y%m%d")
        
        # Open app and enter Password
        run_app_and_login(app_path, app_name)

        # --- Inside The Apps ---
        _open_sales_analysis()
        _set_date_and_options(input_day_str)
        
        # Saving the necessary values
        pyautogui.press("F8")
        time.sleep(1)
        pyautogui.typewrite(str("total_sales"), interval = 0.05)
        pyautogui.press("Enter")
        pyautogui.press('left')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        pyautogui.press('enter')
        time.sleep(0.5)
        
        # Open Sales
        pyautogui.moveTo(*get_tuple("COORDS", "open_total_sales"), duration=0.5)
        time.sleep(0.5)
        pyautogui.click()
        pyautogui.click()
        time.sleep(0.5)
        
        # Save Floors
        for floor_name, coord_key  in [
            ("GF", "floor_GF"),
            ("1F", "floor_1F"),
            ("2F", "floor_2F"),
            ("3F", "floor_3F"),
        ]:
            _save_sales_report(floor_name, coord_key)
        
        # MenuOpn
        pyautogui.moveTo(*get_tuple("COORDS", "menu_open"), duration=0.1)
        time.sleep(0.5)
        pyautogui.click()
        
        # Sales Report
        _daily_sales_analysis(input_day_str)
        
        # Writing in Excel function
        sato.input_sales(input_date.day)
        
        # Popup done
        ctypes.windll.user32.MessageBoxW(
            0,  # HWND — 0 means no owner window
            "Script completed successfully!",
            "Done",
            0x00000040 | 0x00040000 | 0x00010000 # MB_ICONINFORMATION + MB_TOPMOST + MB_SETFOREGROUND
        )
    
        logger.info("=== Sales automation finished successfully ===")

    except Exception as e:
        logger.exception("Automation failed with an error")
        raise
    
    finally:
        end_time = time.time()
        elapsed = end_time - start_time
        mins, secs = divmod(int(elapsed), 60)
        logger.info(f"Start time : {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(start_time))}")
        logger.info(f"End time   : {time.strftime('%Y-%m-%d %H:%M:%S', time.localtime(end_time))}")
        logger.info(f"Elapsed    : {mins}m {secs}s")

if __name__ == "__main__":
    run_sales_automation()