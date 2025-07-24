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
    for attempt in range(60):
        try:
            dlg = Desktop(backend="uia").window(title_re=".*Import Preferences.*")
            if dlg.exists(timeout=1):
                print(f"🖼️ Attempt {attempt+1}: Dialog found.")

                desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
                full_path = os.path.join(desktop_path, "sql_developer_full_desktop.png")
                focused_path = os.path.join(desktop_path, "sql_developer_import_prompt.png")

                pyautogui.screenshot().save(full_path)
                print(f"💾 Full screenshot saved: {full_path}")

                dlg.capture_as_image().save(focused_path)
                print(f"💾 Focused screenshot saved: {focused_path}")

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

def wait_and_dismiss_usage_tracking():
    print("⏳ Waiting for 'Oracle Usage Tracking' dialog...")
    for attempt in range(60):  # wait up to 60 seconds for window to appear
        try:
            dlg = Desktop(backend="win32").window(title="Oracle Usage Tracking", class_name="SunAwtDialog")
            if dlg.exists(timeout=1):
                print(f"🪟 Found 'Oracle Usage Tracking' on attempt {attempt+1}")

                # Screenshot when window is first detected
                desktop_path = os.path.join(os.environ["USERPROFILE"], "Desktop")
                early_path = os.path.join(desktop_path, "oracle_usage_tracking_prompt_early.png")
                dlg.capture_as_image().save(early_path)
                print(f"💾 Screenshot saved: {early_path}")

                # Wait 30 seconds for dialog to fully render
                print("⏳ Waiting 30 seconds for dialog to finish rendering...")
                time.sleep(30)

                # Screenshot before sending keys
                final_path = os.path.join(desktop_path, "oracle_usage_tracking_prompt_before_keys.png")
                dlg.capture_as_image().save(final_path)
                print(f"💾 Screenshot before keys saved: {final_path}")

                # Focus, move mouse, and click to enforce focus in Java dialog
                dlg.set_focus()
                rect = dlg.rectangle()
                center_x = (rect.left + rect.right) // 2
                center_y = (rect.top + rect.bottom) // 2
                pyautogui.moveTo(center_x, center_y)
                time.sleep(0.5)
                pyautogui.click()  # Required to ensure Java dialog receives focus
                time.sleep(0.5)

                # Send Tab, Tab, Enter
                print("⌨️ Sending TAB, TAB, ENTER via pyautogui...")
                pyautogui.press('tab')
                pyautogui.press('tab')
                pyautogui.press('enter')
                print("✅ Oracle Usage Tracking dialog dismissed.")
                return
        except Exception as e:
            print(f"⚠️ Exception: {e}")
        time.sleep(1)

    print("❌ Oracle Usage Tracking dialog not found within timeout.")

def close_explorer_window(window):
    print("❎ Closing File Explorer using .close()...")
    try:
        window.close()
        print("✅ File Explorer closed.")
    except Exception as e:
        print(f"❌ Failed to close File Explorer: {e}")

def close_sql_developer():
    print("❎ Attempting to close SQL Developer window...")
    try:
        dlg = Desktop(backend="win32").window(title="Oracle SQL Developer", class_name="SunAwtFrame")
        if dlg.exists(timeout=5):
            dlg.close()
            print("✅ SQL Developer closed.")
        else:
            print("❌ SQL Developer window not found.")
    except Exception as e:
        print(f"⚠️ Failed to close SQL Developer: {e}")

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
    wait_and_dismiss_usage_tracking()
    close_sql_developer()
    close_explorer_window(explorer_win)

if __name__ == "__main__":
    main()
