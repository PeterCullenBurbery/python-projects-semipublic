import os
import time
import subprocess
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
import pyautogui

def open_shell_desktop():
    print("📂 Opening File Explorer at 'shell:Desktop'...")
    subprocess.Popen(['explorer', 'shell:Desktop'])
    time.sleep(2)

def wait_for_explorer_window():
    print("🪟 Waiting for File Explorer window...")
    for _ in range(10):
        try:
            for w in Desktop(backend="uia").windows(class_name="CabinetWClass"):
                if "desktop" in w.window_text().lower():
                    return w
        except:
            pass
        time.sleep(1)
    return None

def click_sql_developer(explorer_win):
    print("🖱️ Searching for 'SQL Developer' icon...")
    items_view = None
    for ctrl in explorer_win.descendants():
        if ctrl.element_info.control_type == "List" and ctrl.element_info.name == "Items View":
            items_view = ctrl
            break

    if not items_view:
        print("❌ Could not find 'Items View'.")
        return

    for item in items_view.children():
        if item.element_info.name.lower() == "sql developer":
            print("✅ Found 'SQL Developer'. Double-clicking...")
            item.double_click_input()
            return

    print("❌ 'SQL Developer' icon not found.")

def wait_and_handle_import_prompt():
    print("⏳ Waiting for 'Confirm Import Preferences' dialog...")
    for attempt in range(30):  # up to 30 seconds
        try:
            dlg = Desktop(backend="uia").window(title_re=".*Import Preferences.*")
            if dlg.exists(timeout=1):
                print(f"🖼️ Attempt {attempt+1}: Dialog found.")

                # Screenshot paths
                desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
                full_path = os.path.join(desktop_path, "sql_developer_full_desktop.png")
                focused_path = os.path.join(desktop_path, "sql_developer_import_prompt.png")

                # 🌐 Full screenshot with pyautogui
                full_screenshot = pyautogui.screenshot()
                full_screenshot.save(full_path)
                print(f"💾 Full screenshot saved: {full_path}")

                # 🎯 Focused screenshot with pywinauto (no background)
                focused_image = dlg.capture_as_image()
                focused_image.save(focused_path)
                print(f"💾 Focused screenshot saved: {focused_path}")

                # 🚪 Dismiss dialog
                print("👉 Sending Alt+N to dismiss dialog...")
                dlg.set_focus()
                time.sleep(0.5)
                send_keys('%N')
                print("✅ Sent Alt+N.")
                return
        except Exception as e:
            print(f"⚠️ Exception: {e}")
        time.sleep(1)

    print("❌ Dialog not found or failed to handle within timeout.")

def main():
    open_shell_desktop()
    explorer_win = wait_for_explorer_window()
    if not explorer_win:
        print("❌ Could not find Desktop File Explorer window.")
        return

    explorer_win.set_focus()
    time.sleep(1)
    click_sql_developer(explorer_win)
    wait_and_handle_import_prompt()

if __name__ == "__main__":
    main()