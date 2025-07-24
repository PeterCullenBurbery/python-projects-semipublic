import subprocess
import time
from pywinauto import Desktop

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

def list_desktop_items(explorer_win):
    print("📋 Listing desktop items...")
    items_view = None
    for ctrl in explorer_win.descendants():
        if ctrl.element_info.control_type == "List" and ctrl.element_info.name == "Items View":
            items_view = ctrl
            break

    if not items_view:
        print("❌ Could not find 'Items View'.")
        return

    items = items_view.children()
    for i, item in enumerate(items, start=1):
        print(f"{i}. {item.element_info.name}")

def main():
    open_shell_desktop()
    explorer_win = wait_for_explorer_window()
    if not explorer_win:
        print("❌ Could not find Desktop File Explorer window.")
        return

    explorer_win.set_focus()
    time.sleep(1)
    list_desktop_items(explorer_win)

if __name__ == "__main__":
    main()
