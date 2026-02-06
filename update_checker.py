import sys
import os
import json
import tempfile
import subprocess
from urllib.request import urlopen, Request
from urllib.error import URLError
import ctypes

GITHUB_REPO = "nathannncurtis/Study-Aggregator"

def get_current_version():
    """Read version from version.txt bundled alongside the executable."""
    try:
        if getattr(sys, 'frozen', False):
            base_path = os.path.dirname(sys.executable)
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base_path, 'version.txt'), 'r') as f:
            return f.read().strip()
    except Exception:
        return "0.0.0"

CURRENT_VERSION = get_current_version()
GITHUB_API_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"


def parse_version(version_str):
    version_str = version_str.lstrip("v").strip()
    try:
        return tuple(int(x) for x in version_str.split("."))
    except (ValueError, AttributeError):
        return (0, 0, 0)


def check_for_update():
    """Check GitHub for a newer release. Returns (tag, asset_url, release_page_url) or None."""
    try:
        req = Request(GITHUB_API_URL, headers={"User-Agent": "Study-Aggregator-UpdateChecker"})
        with urlopen(req, timeout=15) as response:
            data = json.loads(response.read().decode("utf-8"))

        latest_tag = data.get("tag_name", "")
        latest_version = parse_version(latest_tag)
        current_version = parse_version(CURRENT_VERSION)

        if latest_version <= current_version:
            return None

        release_page_url = data.get("html_url", "")

        # Find installer asset (.exe) in release assets
        asset_url = None
        for asset in data.get("assets", []):
            if asset.get("name", "").lower().endswith(".exe"):
                asset_url = asset.get("browser_download_url")
                break

        return (latest_tag, asset_url, release_page_url)
    except Exception:
        return None


def message_box(title, message, style=0x40):
    """Show a Windows message box. Returns button ID."""
    # 0x40 = MB_ICONINFORMATION, 0x44 = MB_YESNO | MB_ICONINFORMATION
    return ctypes.windll.user32.MessageBoxW(0, message, title, style)


def download_and_install(asset_url, release_page_url):
    """Download installer and run it. Falls back to opening browser."""
    if asset_url:
        try:
            req = Request(asset_url, headers={"User-Agent": "Study-Aggregator-UpdateChecker"})
            with urlopen(req, timeout=300) as response:
                installer_data = response.read()

            installer_path = os.path.join(tempfile.gettempdir(), "StudyAggregatorSetup.exe")
            with open(installer_path, "wb") as f:
                f.write(installer_data)

            subprocess.Popen([installer_path], shell=False)
            return
        except Exception:
            pass

    # Fallback: open release page in browser
    if release_page_url:
        try:
            os.startfile(release_page_url)
            return
        except Exception:
            pass

    message_box(
        "Study Aggregator Update",
        "Failed to download the update automatically.\n\n"
        f"Please visit https://github.com/{GITHUB_REPO}/releases to download manually.",
    )


def main():
    result = check_for_update()
    if result is None:
        sys.exit(0)

    tag, asset_url, release_page_url = result

    # MB_YESNO (0x04) | MB_ICONINFORMATION (0x40) = 0x44
    response = message_box(
        "Study Aggregator Update",
        f"A new version ({tag}) of Study Aggregator is available.\n\n"
        "Would you like to download and install it now?",
        0x44,
    )

    # IDYES = 6
    if response == 6:
        download_and_install(asset_url, release_page_url)


if __name__ == "__main__":
    main()
