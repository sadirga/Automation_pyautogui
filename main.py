import pyautogui
import time
import sales_auto as sato
from datetime import date, timedelta, datetime
import ctypes
import os
import tkinter as tk
from tkinter import simpledialog


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
        
def run_sales_automation():
    password = get_password()
        
    # --- Show Desktop ---
    pyautogui.hotkey('win', 'd')  # Press Windows + D
    time.sleep(1)
    
    # Opening the targeted app
    pyautogui.moveTo(114, 25, duration=0.5)
    time.sleep(1)
    pyautogui.click()
    pyautogui.press('enter')

    # Short delay before typing
    time.sleep(6)

    # Enter Password
    pyautogui.typewrite(password, interval=0.05)  # interval = delay between keystrokes
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')

    # --- Inside The Apps ---
    _open_sales_analysis()
    _set_date_and_options()
    
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
    pyautogui.moveTo(387, 278, duration=0.5)
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
        
    sato.input_sales()
   
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
        pyautogui.moveTo(x, y, duration=0.3)
        time.sleep(0.5)
        pyautogui.click()
        pyautogui.click()
        time.sleep(1)
    
def _set_date_and_options():
    # Setting yesterday date as the targeted date
    yesterday_str = (datetime.now() - timedelta(days=1)).strftime("%Y%m%d")
    pyautogui.typewrite(yesterday_str, interval=0.08)
    time.sleep(1)
    pyautogui.moveTo(1664, 113, duration=0.3)
    time.sleep(1)
    pyautogui.click()

    # Option Button
    pyautogui.moveTo(1844, 103, duration=0.5)
    time.sleep(1)
    pyautogui.click()

    # Unit Rp
    pyautogui.moveTo(849, 525, duration=0.5)
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