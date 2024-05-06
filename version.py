from win32com.client import Dispatch
import os

def get_version_via_com(filename):
    parser = Dispatch("Scripting.FileSystemObject")
    try:
        version = parser.GetFileVersion(filename)
    except Exception:
        return None
    return version

def identify_browser_versions():
    browsers = {
        "Chrome": [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
        ],
        "Firefox": [
            r"C:\Program Files\Mozilla Firefox\firefox.exe",
            r"C:\Program Files (x86)\Mozilla Firefox\firefox.exe"
        ],
        "Edge": [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        ],
        "Brave": [
            r"C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe",
            r"C:\Program Files (x86)\BraveSoftware\Brave-Browser\Application\brave.exe",
        ],
        "Opera": [
            r"C:\Program Files\Opera\launcher.exe",
            r"C:\Program Files (x86)\Opera\launcher.exe",
        ],
        # Add more browsers as needed
    }

    browser_versions = {}

    for browser, paths in browsers.items():
        for path in paths:
            if os.path.exists(path):
                version = get_version_via_com(path)
                if version:
                    browser_versions[browser] = version
                break  # Stop searching if the browser is found

    return browser_versions

if __name__ == "__main__":
    all_browser_versions = identify_browser_versions()

    if all_browser_versions:
        for browser, version in all_browser_versions.items():
            print(f"{browser} version: {version}")
    else:
        print("No supported browsers found.")
