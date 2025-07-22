import time
from pywinauto import Desktop, Application

def wait_and_dump_window(title, timeout=15):
    print(f"⌛ Waiting for window titled: '{title}'...")
    for _ in range(timeout):
        try:
            win = Desktop(backend="uia").window(title=title)
            if win.exists() and win.is_visible():
                print("✅ Found the installer window!")
                return win
        except:
            pass
        time.sleep(1)
    raise TimeoutError("❌ Installer window not found within timeout.")

def explore_window_controls(window):
    print("\n🧩 Dumping control tree:")
    window.print_control_identifiers()

    print("\n📋 Summary of children and properties:")
    for ctrl in window.descendants():
        try:
            print(f"\n=== {ctrl.window_text()} ===")
            print(f"Control Type: {ctrl.control_type()}")
            print(f"Automation ID: {ctrl.element_info.automation_id}")
            print(f"Class Name: {ctrl.element_info.class_name}")
            print(f"Framework: {ctrl.element_info.framework_id}")
            print(f"Handle: {ctrl.handle}")
            print(f"Enabled: {ctrl.is_enabled()}")
            print(f"Visible: {ctrl.is_visible()}")
            print(f"Rectangle: {ctrl.rectangle()}")
        except Exception as e:
            print(f"(⚠️ Skipped one element due to error: {e})")

if __name__ == "__main__":
    title = "Python 3.13.5 (64-bit) Setup"
    win = wait_and_dump_window(title)
    explore_window_controls(win)
