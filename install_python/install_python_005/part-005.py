import time
from pywinauto.application import Application
from pywinauto import Desktop

def wait_for_optional_features(title="Python 3.13.5 (64-bit) Setup", timeout=15):
    print(f"⌛ Waiting for '{title}' window (Optional Features)...")
    for _ in range(timeout):
        try:
            win = Desktop(backend="uia").window(title=title)
            if win.exists() and win.is_visible():
                print("✅ Found the Optional Features window!")
                return win
        except:
            pass
        time.sleep(1)
    raise TimeoutError(f"❌ '{title}' window not found.")

def check_feature(dlg, title):
    try:
        box = dlg.child_window(title=title, control_type="CheckBox")
        if not box.get_toggle_state():
            box.toggle()
        print(f"☑️ Enabled: {title}")
    except Exception as e:
        print(f"⚠️ Could not enable '{title}': {e}")

def automate_optional_features():
    dlg = wait_for_optional_features()
    app = Application(backend="uia").connect(title=dlg.window_text())
    dlg = app.window(title=dlg.window_text())

    features = [
        "Documentation",
        "pip",
        "tcl/tk and IDLE",
        "Python test suite",
        "py launcher",
        "for all users (requires admin privileges)"
    ]

    for feature in features:
        check_feature(dlg, feature)

    # Click the Next button
    try:
        dlg.child_window(title="Next", control_type="Button").invoke()
        print("➡️ Clicked 'Next'")
    except Exception as e:
        print(f"⚠️ Could not click 'Next': {e}")

if __name__ == "__main__":
    automate_optional_features()
