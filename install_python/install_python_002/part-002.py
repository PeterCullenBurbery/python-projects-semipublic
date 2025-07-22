import time
from pywinauto.application import Application
from pywinauto import Desktop

def wait_for_installer(title, timeout=15):
    print(f"⌛ Waiting for '{title}' window...")
    for _ in range(timeout):
        try:
            return Desktop(backend="uia").window(title=title)
        except:
            time.sleep(1)
    raise RuntimeError(f"❌ Timeout waiting for '{title}' window.")

def automate_python_installer():
    title = "Python 3.13.5 (64-bit) Setup"
    window = wait_for_installer(title)
    app = Application(backend="uia").connect(title=title)
    dlg = app.window(title=title)

    print("✅ Installer window found")

    try:
        # Find all descendants
        all_controls = dlg.descendants()

        for ctrl in all_controls:
            name = ctrl.window_text()
            if name == "Use admin privileges when installing py.exe":
                ctrl.invoke()
                print("☑️ Clicked 'Use admin privileges'")
            elif name == "Add python.exe to PATH":
                ctrl.invoke()
                print("☑️ Clicked 'Add python.exe to PATH'")
            elif name == "Customize installation":
                ctrl.invoke()
                print("➡️ Clicked 'Customize installation'")
    except Exception as e:
        print(f"⚠️ Error while interacting: {e}")

if __name__ == "__main__":
    automate_python_installer()
