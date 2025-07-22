import os
import tempfile
import urllib.request
import ssl
import certifi
import shutil

def download_python_installer():
    url = "https://www.python.org/ftp/python/3.13.5/python-3.13.5-amd64.exe"
    temp_dir = tempfile.gettempdir()
    output_path = os.path.join(temp_dir, "python-3.13.5-amd64.exe")

    print(f"Downloading to: {output_path}")
    try:
        ssl_context = ssl.create_default_context(cafile=certifi.where())
        with urllib.request.urlopen(url, context=ssl_context) as response, open(output_path, 'wb') as out_file:
            shutil.copyfileobj(response, out_file)
        print("Download complete.")
    except Exception as e:
        print(f"Download failed: {e}")

if __name__ == "__main__":
    download_python_installer()
