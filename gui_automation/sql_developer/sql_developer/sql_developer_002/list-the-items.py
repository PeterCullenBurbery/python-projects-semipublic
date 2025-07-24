import subprocess
import time
from pywinauto import Desktop
from pywinauto.keyboard import send_keys

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

def main():
    open_shell_desktop()
    explorer_win = wait_for_explorer_window()
    if not explorer_win:
        print("❌ Could not find Desktop File Explorer window.")
        return

    explorer_win.set_focus()
    time.sleep(1)
    click_sql_developer(explorer_win)

if __name__ == "__main__":
    main()