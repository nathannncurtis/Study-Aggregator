import sys
import os
import shutil
import subprocess
import tempfile

from windows_toasts import Toast, ToastDisplayImage, InteractableToastNotifier, ToastButton

# ── Configuration ──────────────────────────────────────────────────────────────
UPDATE_SHARE = r"\\SERVER\SHARE\StudyAggregator"  # TODO: set actual UNC path
AUMID = "Ronsin.StudyAggregator"
INSTALLER_NAME = "StudyAggregatorSetup.exe"
BLOCKING_PROCESSES = {"study aggregator.exe", "7z.exe", "7za.exe"}


def get_local_version():
    try:
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
        with open(os.path.join(base, "version.txt")) as f:
            return f.read().strip()
    except Exception:
        return "0.0.0"


def get_remote_version():
    try:
        with open(os.path.join(UPDATE_SHARE, "version.txt")) as f:
            return f.read().strip()
    except Exception:
        return None


def parse_version(v):
    try:
        return tuple(int(x) for x in v.strip().lstrip("v").split("."))
    except Exception:
        return (0, 0, 0)


def is_process_running():
    try:
        result = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=10,
        )
        for line in result.stdout.lower().splitlines():
            for proc in BLOCKING_PROCESSES:
                if proc in line:
                    return True
    except Exception:
        pass
    return False


def install_update():
    try:
        src = os.path.join(UPDATE_SHARE, INSTALLER_NAME)
        dst = os.path.join(tempfile.gettempdir(), INSTALLER_NAME)
        shutil.copy2(src, dst)
        subprocess.Popen([dst], shell=False)
    except Exception:
        pass


def on_button_clicked(args):
    if args.arguments == "update":
        install_update()
    sys.exit(0)


def on_dismissed(args):
    sys.exit(0)


def main():
    local_ver = get_local_version()
    remote_ver = get_remote_version()

    if remote_ver is None:
        sys.exit(0)

    if parse_version(remote_ver) <= parse_version(local_ver):
        sys.exit(0)

    if is_process_running():
        sys.exit(0)

    toast = Toast()
    toast.text_fields = [
        "Study Aggregator Update Available",
        f"Version {remote_ver} is available (installed: {local_ver})",
    ]

    toast.AddAction(ToastButton("Update Now", arguments="update"))
    toast.AddAction(ToastButton("Dismiss", arguments="dismiss"))

    toast.on_activated = on_button_clicked
    toast.on_dismissed = on_dismissed

    notifier = InteractableToastNotifier(AUMID)
    notifier.show_toast(toast)


if __name__ == "__main__":
    main()
