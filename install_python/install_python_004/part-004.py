import time
from pywinauto.application import Application
from pywinauto import Desktop

def wait_for_installer(title, timeout=15):
    print(f"⌛ Waiting for '{title}' window...")
    for _ in range(timeout):
        try:
            win = Desktop(backend="uia").window(title=title)
            if win.exists() and win.is_visible():
                print("✅ Found the installer window!")
                return win
        except:
            pass
        time.sleep(1)
    raise TimeoutError(f"❌ Installer window not found.")

def automate_python_installer():
    title = "Python 3.13.5 (64-bit) Setup"
    window = wait_for_installer(title)
    app = Application(backend="uia").connect(title=title)
    dlg = app.window(title=title)

    try:
        dlg.child_window(title="Use admin privileges when installing py.exe", control_type="CheckBox").invoke()
        print("☑️ Clicked 'Use admin privileges'")
    except Exception as e:
        print(f"⚠️ Couldn't click admin checkbox: {e}")

    try:
        dlg.child_window(title="Add python.exe to PATH", control_type="CheckBox").invoke()
        print("☑️ Clicked 'Add python.exe to PATH'")
    except Exception as e:
        print(f"⚠️ Couldn't click PATH checkbox: {e}")

    try:
        dlg.child_window(title="Customize installation", control_type="Button").invoke()
        print("➡️ Clicked 'Customize installation'")
    except Exception as e:
        print(f"⚠️ Couldn't click 'Customize installation': {e}")

if __name__ == "__main__":
    automate_python_installer()
