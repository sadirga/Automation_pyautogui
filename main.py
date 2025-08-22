import pyautogui
import psutil
import time
import sales_auto as sato
from datetime import date, timedelta, datetime
import ctypes
import os
import tkinter as tk
from tkinter import simpledialog
import configparser

def ask_for_date():
    root = tk.Tk()
    root.withdraw()  # hide main window
    day_str = simpledialog.askstring("Input", "Enter day number (1-31):")
    root.destroy()
    today = datetime.today()
    return datetime(today.year, today.month, int(day_str))
    
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

def run_app_and_login(shortcut_path, proc_name):
    # Start via shortcut
    os.startfile(shortcut_path)

    # Wait for process showing
    proc = wait_for_process(proc_name, timeout=20)

    # Wait for UI to show
    wait_for_pixel(1256, 622, (255,255,255)) # Need to adjust the pixel on other device

    # Enter password
    password = get_password()
    pyautogui.typewrite(password, interval=0.05)
    
    pyautogui.press("enter")
    time.sleep(0.5)
    pyautogui.press("enter")
        
def run_sales_automation():
    
    # Press Windows + D
    pyautogui.hotkey("winleft", "d")
    time.sleep(2)
    
    # Load config
    config = configparser.ConfigParser()
    config.read("config.ini")
    
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
    pyautogui.typewrite(str("totalSales"), interval = 0.05)
    pyautogui.press("Enter")
    pyautogui.press('left')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(3)
    
    # Open Sales
    pyautogui.moveTo(387, 278, duration=0.5) # 1st time enter
    time.sleep(1)
    pyautogui.click()
    pyautogui.click()
    time.sleep(1)
    
    # Save Floors
    for floor_name, coords in [
        ("GF", (398, 544)),
        ("1F", (380, 290)),
        ("2F", (395, 371)),
        ("3F", (395, 443))
    ]:
        _save_sales_report(floor_name, coords)
    
    # MenuOpn
    pyautogui.moveTo(191, 32, duration=0.1)
    time.sleep(1)
    pyautogui.click()
    
    # Sales Report
    _daily_sales_analysis(input_day_str)
    
    # Writing in Excel function
    day_number = input_date.day
    sato.input_sales(day_number)
    
    
    
    # Popup done
    ctypes.windll.user32.MessageBoxW(
        0,  # HWND — 0 means no owner window
        "Script completed successfully!",
        "Done",
        0x00000040 | 0x00040000 | 0x00010000 # MB_ICONINFORMATION + MB_TOPMOST + MB_SETFOREGROUND
    )
    
def _open_sales_analysis():
    
    steps = [
        (105, 115),  # Store Sales Analysis
        (105, 133),  # Sales News
        (105, 150),  # Total Sales News
    ]
    for x, y in steps:
        pyautogui.moveTo(x, y, duration=0.2)
        time.sleep(0.5)
        pyautogui.click()
        pyautogui.click()
        time.sleep(1)
        
def _daily_sales_analysis(input_day):
    
    steps = [
        (97, 186),  # Daily Sales analysis
        (97, 240),  # Sales Report
    ]
    for x, y in steps:
        pyautogui.moveTo(x, y, duration=0.3)
        time.sleep(0.5)
        pyautogui.click()
        pyautogui.click()
        time.sleep(1)
    
    # Set Date
    pyautogui.moveTo(354, 109, duration=0.2) # cursor mengarah pilihan date/tanggal
    pyautogui.click()
    input_day_str = input_day
    time.sleep(0.5)
    pyautogui.typewrite(input_day_str, interval=0.08)
    pyautogui.typewrite(input_day_str, interval=0.08)
    
    # Open feature
    pyautogui.moveTo(141, 178, duration=0.3) # Cursor memilih tab "Pay Sumamry"
    pyautogui.click()
    time.sleep(1)
    pyautogui.press('f2')
    wait_for_pixel(451,228, (0,120,215)) # Wait until loading data finish
    
    # Save File
    pyautogui.press('f8')
    pyautogui.typewrite("allbrand", interval=0.05)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('left')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)
    
def _set_date_and_options(input_day_str):
    # Setting input_day_str date as the targeted date
    pyautogui.typewrite(input_day_str, interval=0.08)
    time.sleep(1)
    pyautogui.moveTo(1664, 113, duration=0.3) # koordinat pilihan tanggal
    time.sleep(1)
    pyautogui.click()

    # Option Button
    pyautogui.moveTo(1844, 103, duration=0.5) # koordinat pilihan option
    time.sleep(1)
    pyautogui.click()

    # Unit Rp
    pyautogui.moveTo(849, 525, duration=0.5) # koordinat pilihan unit
    time.sleep(1)
    pyautogui.click()

    # OK Button
    pyautogui.moveTo(1064, 474, duration=0.5)
    time.sleep(1)
    pyautogui.click()

def _save_sales_report(floor_name, coords):

    # Floor Selection
    pyautogui.moveTo(*coords, duration=0.5)
    time.sleep(1)
    pyautogui.click()
    pyautogui.click()
    time.sleep(0.7)

    # Save File
    pyautogui.press('F8')
    time.sleep(1)
    pyautogui.typewrite(floor_name, interval=0.05)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('left')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(0.5)
    pyautogui.press('enter')
    time.sleep(1)

    # Go Back
    pyautogui.moveTo(373, 245, duration=0.5)
    time.sleep(1)
    pyautogui.click()
    pyautogui.click()
   
    
    
if __name__ == "__main__":
    run_sales_automation()