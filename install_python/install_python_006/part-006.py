import time
from pywinauto.application import Application
from pywinauto import Desktop

def wait_for_advanced_options(title="Python 3.13.5 (64-bit) Setup", timeout=15):
    print(f"⌛ Waiting for '{title}' window (Advanced Options)...")
    for _ in range(timeout):
        try:
            win = Desktop(backend="uia").window(title=title)
            if win.exists() and win.is_visible():
                print("✅ Found the Advanced Options window!")
                return win
        except:
            pass
        time.sleep(1)
    raise TimeoutError(f"❌ '{title}' window not found.")

def check_feature(dlg, title):
    try:
        checkbox = dlg.child_window(title=title, control_type="CheckBox")
        if not checkbox.get_toggle_state():
            checkbox.toggle()
        print(f"☑️ Enabled: {title}")
    except Exception as e:
        print(f"⚠️ Could not enable '{title}': {e}")

def configure_advanced_options():
    dlg = wait_for_advanced_options()
    app = Application(backend="uia").connect(title=dlg.window_text())
    dlg = app.window(title=dlg.window_text())

    features = [
        "Install Python 3.13 for all users",
        "Associate files with Python (requires the 'py' launcher)",
        "Create shortcuts for installed applications",
        "Add Python to environment variables",
        "Precompile standard library"
    ]

    for feature in features:
        check_feature(dlg, feature)

    print("✅ Configuration complete. Install button not clicked.")

if __name__ == "__main__":
    configure_advanced_options()
