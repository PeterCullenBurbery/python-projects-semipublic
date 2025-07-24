import os
import time
import subprocess
import ctypes
import psutil
from pywinauto import Desktop
from pywinauto.keyboard import send_keys
import pyautogui

def block_input():
    if ctypes.windll.user32.BlockInput(True):
        print("🔒 Input blocked.")
        return True
    else:
        print("❌ Failed to block input. Are you running as Administrator?")
        return False

def unblock_input():
    ctypes.windll.user32.BlockInput(False)
    print("🔓 Input unblocked.")

def get_sqldeveloper_pids():
    pids = set()
    for proc in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            cmd = ' '.join(proc.info.get('cmdline') or [])
            if 'sqldeveloper' in cmd.lower():
                pids.add(proc.info['pid'])
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids

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
        if item.element_info.name.strip().lower() == "sql developer":
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

def close_windows_by_pid(pid):
    print(f"❎ Attempting to close SQL Developer windows for PID: {pid}")
    found = False
    for w in Desktop(backend="win32").windows():
        try:
            if w.process_id() == pid:
                print(f"🔍 Closing window: '{w.window_text()}' [Class: {w.class_name()}]")
                w.close()
                found = True
        except Exception as e:
            print(f"⚠️ Failed to close window: {e}")
    if found:
        print("✅ SQL Developer window(s) closed.")
    else:
        print("❌ No windows found matching the PID.")

def close_explorer_window(window):
    print("❎ Closing File Explorer using .close()...")
    try:
        window.close()
        print("✅ File Explorer closed.")
    except Exception as e:
        print(f"❌ Failed to close File Explorer: {e}")

def terminate_launcher_if_idle(pids):
    print("🧼 Checking if launcher processes are still running...")
    for pid in pids:
        try:
            proc = psutil.Process(pid)
            name = proc.name().lower()
            cmd = ' '.join(proc.cmdline()).lower()
            if 'sqldeveloper' in cmd and 'java' not in name:
                print(f"💀 Terminating leftover launcher PID: {pid} ({name})")
                proc.terminate()
                proc.wait(timeout=5)
                print("✅ Launcher process terminated.")
        except Exception as e:
            print(f"⚠️ Could not terminate PID {pid}: {e}")

def main():
    input_was_blocked = block_input()

    try:
        print("🔍 Capturing SQL Developer PIDs before launch...")
        before_pids = get_sqldeveloper_pids()
        print(f"📋 Before PIDs: {before_pids}")

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

        print("⏳ Waiting 15 seconds before shutdown...")
        time.sleep(15)

        print("🔍 Capturing SQL Developer PIDs after launch...")
        after_pids = get_sqldeveloper_pids()
        print(f"📋 After PIDs: {after_pids}")

        new_pids = after_pids - before_pids
        print(f"🆕 New SQL Developer PIDs: {new_pids}")

        for pid in new_pids:
            close_windows_by_pid(pid)

        close_explorer_window(explorer_win)
        terminate_launcher_if_idle(new_pids)

    finally:
        if input_was_blocked:
            unblock_input()

if __name__ == "__main__":
    main()
