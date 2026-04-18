import pyautogui
import pyperclip
import time
import os

def extract_via_ui():
    print("Opening PDF...")
    # Open the PDF using the default system viewer (Foxit/WPS/Edge)
    os.startfile(r"e:\教学管理\信息系统项目管理师教程-第四版.pdf")
    
    print("Waiting for PDF viewer to load (10 seconds)...")
    time.sleep(10)
    
    print("Attempting to jump to page 349 (CTRL+G typically or typing page number)...")
    # Most readers respond to CTRL+G or CTRL+Shift+N for page jump
    pyautogui.hotkey('ctrl', 'g')
    time.sleep(1)
    pyautogui.write('349')
    pyautogui.press('enter')
    
    print("Waiting for page to render (3 seconds)...")
    time.sleep(3)
    
    print("Selecting all text on current view (CTRL+A) and Copying (CTRL+C)...")
    # Click somewhere safe in the middle to focus
    pyautogui.click(960, 540)
    time.sleep(0.5)
    pyautogui.hotkey('ctrl', 'a')
    time.sleep(1)
    pyautogui.hotkey('ctrl', 'c')
    time.sleep(1)
    
    text = pyperclip.paste()
    print(f"Extracted {len(text)} characters from clipboard.")
    
    out_path = r"e:\教学管理\ch11_ui_extracted.txt"
    with open(out_path, "w", encoding="utf-8") as f:
        f.write(text)
        
    print(f"Saved to {out_path}. Please check output.")

if __name__ == "__main__":
    extract_via_ui()
