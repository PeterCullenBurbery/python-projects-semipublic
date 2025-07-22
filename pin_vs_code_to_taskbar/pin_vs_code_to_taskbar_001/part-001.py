import subprocess
import time
import ctypes
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

def close_explorer_window(window):
    print("❎ Closing File Explorer using .close()...")
    try:
        window.close()
        print("✅ File Explorer closed.")
    except Exception as e:
        print(f"❌ Failed to close File Explorer: {e}")

def main():
    input_was_blocked = block_input()

    try:
        open_shell_desktop()
        explorer_win = wait_for_explorer_window()
        if not explorer_win:
            print("❌ Could not find Desktop File Explorer window.")
            return

        explorer_win.set_focus()
        time.sleep(1)

        print("🔍 Locating 'Items View'...")
        items_view = None
        for ctrl in explorer_win.descendants():
            if ctrl.element_info.control_type == "List" and ctrl.element_info.name == "Items View":
                items_view = ctrl
                break

        if not items_view:
            print("❌ Could not find 'Items View'.")
            return

        print("🔍 Searching for item 'VSCode'...")
        target_item = None
        for item in items_view.children():
            if item.element_info.name.strip().lower() == "vscode":
                target_item = item
                break

        if not target_item:
            print("❌ 'VSCode' not found.")
            return

        print("🖱️ Right-clicking 'VSCode'...")
        try:
            target_item.set_focus()
            target_item.right_click_input()
        except Exception as e:
            print(f"⚠️ Error right-clicking: {e}")
            return

        print("⏳ Waiting for context menu...")
        time.sleep(1)

        print("📌 Checking pin status in context menu...")
        try:
            menu = Desktop(backend="uia").window(control_type="Menu")
            pin_item = None
            unpin_item = None

            for item in menu.descendants(control_type="MenuItem"):
                name = item.element_info.name.strip().lower()
                if name == "pin to taskbar":
                    pin_item = item
                elif name == "unpin from taskbar":
                    unpin_item = item

            if unpin_item:
                print("🔒 Already pinned to taskbar. Skipping.")
            elif pin_item:
                print("📌 Found 'Pin to taskbar'. Clicking...")
                pin_item.click_input()
                print("✅ VSCode pinned to taskbar.")
            else:
                print("❌ Neither 'Pin to taskbar' nor 'Unpin from taskbar' found.")

            print("🎹 Sending Escape to close context menu...")
            send_keys("{ESC}")
            time.sleep(0.5)

        except Exception as e:
            print(f"⚠️ Failed to handle context menu: {e}")

        time.sleep(1)  # Delay to let animations finish
        close_explorer_window(explorer_win)

    finally:
        if input_was_blocked:
            unblock_input()

if __name__ == "__main__":
    main()
