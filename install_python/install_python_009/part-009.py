import os
import tempfile
import urllib.request
import ssl
import certifi
import shutil
import subprocess
import time
from pywinauto import Desktop, Application

INSTALLER_NAME = "python-3.13.5-amd64.exe"
INSTALLER_URL = f"https://www.python.org/ftp/python/3.13.5/{INSTALLER_NAME}"
INSTALLER_PATH = os.path.join(tempfile.gettempdir(), INSTALLER_NAME)

def download_installer():
    print(f"⬇️ Downloading Python installer to: {INSTALLER_PATH}")
    ssl_context = ssl.create_default_context(cafile=certifi.where())
    with urllib.request.urlopen(INSTALLER_URL, context=ssl_context) as response, open(INSTALLER_PATH, 'wb') as out_file:
        shutil.copyfileobj(response, out_file)
    print("✅ Download complete.")

def launch_installer():
    print("🚀 Launching installer...")
    subprocess.Popen([INSTALLER_PATH])
    time.sleep(2)

def wait_for_window(title, timeout=30):
    print(f"⌛ Waiting for '{title}' window...")
    for _ in range(timeout):
        try:
            win = Desktop(backend="uia").window(title=title)
            if win.exists() and win.is_visible():
                print(f"✅ Found: {title}")
                return win
        except:
            pass
        time.sleep(1)
    raise TimeoutError(f"❌ Timeout waiting for '{title}'")

def dump_controls(win):
    print("\n🧩 Dumping controls:")
    win.print_control_identifiers()

def toggle_checkbox(dlg, title):
    try:
        box = dlg.child_window(title=title, control_type="CheckBox")
        if not box.get_toggle_state():
            box.toggle()
        print(f"☑️ Enabled: {title}")
    except Exception as e:
        print(f"⚠️ Could not enable '{title}': {e}")

def part_004_configure_initial_screen(dlg):
    toggle_checkbox(dlg, "Use admin privileges when installing py.exe")
    toggle_checkbox(dlg, "Add python.exe to PATH")
    try:
        dlg.child_window(title="Customize installation", control_type="Button").invoke()
        print("➡️ Clicked 'Customize installation'")
    except Exception as e:
        print(f"⚠️ Could not click Customize installation: {e}")

def part_005_configure_optional_features(dlg):
    options = [
        "Documentation",
        "pip",
        "tcl/tk and IDLE",
        "Python test suite",
        "py launcher",
        "for all users (requires admin privileges)",
    ]
    for opt in options:
        toggle_checkbox(dlg, opt)
    try:
        dlg.child_window(title="Next", control_type="Button").invoke()
        print("➡️ Clicked 'Next'")
    except Exception as e:
        print(f"⚠️ Could not click Next: {e}")

def part_006_configure_advanced_options(dlg):
    options = [
        "Install Python 3.13 for all users",
        "Associate files with Python (requires the 'py' launcher)",
        "Create shortcuts for installed applications",
        "Add Python to environment variables",
        "Precompile standard library"
    ]
    for opt in options:
        toggle_checkbox(dlg, opt)

    # Now click Install
    try:
        dlg.child_window(title="Install", control_type="Button").invoke()
        print("🚀 Clicked 'Install'")
    except Exception as e:
        print(f"⚠️ Could not click Install: {e}")

def find_close_button_excluding_title_bar(dlg):
    """Find the 'Close' button that is not part of the title bar."""
    for ctrl in dlg.descendants(control_type="Button", title="Close"):
        parent = ctrl.parent()
        in_title_bar = False
        while parent:
            try:
                if parent.friendly_class_name() == "TitleBar":
                    in_title_bar = True
                    break
                parent = parent.parent()
            except Exception:
                break
        if not in_title_bar:
            return ctrl
    return None

def wait_for_success_and_close(title="Python 3.13.5 (64-bit) Setup", timeout=120):
    print(f"⌛ Waiting up to {timeout} seconds for setup success message and Close button...")
    for elapsed in range(timeout):
        try:
            win = Desktop(backend="uia").window(title=title)
            if not (win.exists() and win.is_visible()):
                time.sleep(1)
                continue

            app = Application(backend="uia").connect(title=title)
            dlg = app.window(title=title)

            # Check for the success message
            texts = [ctrl.window_text() for ctrl in dlg.descendants() if ctrl.window_text()]
            has_success = any("Setup was successful" in text for text in texts)

            # Look for Close button (excluding title bar)
            close_btn = find_close_button_excluding_title_bar(dlg)
            has_close = close_btn is not None and close_btn.is_enabled()

            if has_success:
                print(f"✅ [{elapsed+1}s] Found success message.")
            if close_btn:
                print(f"✅ [{elapsed+1}s] Found dialog Close button.")
            else:
                print(f"❌ [{elapsed+1}s] Close button not found (excluding title bar).")

            if has_success and has_close:
                # close_btn.invoke()
                # print("🛑 Clicked 'Close'. Installer exited.")
                print("close")
                return
        except Exception as e:
            pass  # You can optionally log e here if debugging
        time.sleep(1)

    print("⚠️ Timeout waiting for success message and Close button.")

def main():
    download_installer()
    launch_installer()

    # Part 004
    win = wait_for_window("Python 3.13.5 (64-bit) Setup")
    app = Application(backend="uia").connect(title=win.window_text())
    dlg = app.window(title=win.window_text())
    dump_controls(dlg)
    part_004_configure_initial_screen(dlg)

    # Part 005
    win = wait_for_window("Python 3.13.5 (64-bit) Setup")
    app = Application(backend="uia").connect(title=win.window_text())
    dlg = app.window(title=win.window_text())
    dump_controls(dlg)
    part_005_configure_optional_features(dlg)

    # Part 006
    win = wait_for_window("Python 3.13.5 (64-bit) Setup")
    app = Application(backend="uia").connect(title=win.window_text())
    dlg = app.window(title=win.window_text())
    dump_controls(dlg)
    part_006_configure_advanced_options(dlg)

    # ✅ New final step
    wait_for_success_and_close()

if __name__ == "__main__":
    main()
