import sys
import os
import shutil
import subprocess
import tempfile
import threading

from windows_toasts import Toast, InteractableWindowsToaster, ToastButton

# ── Configuration ──────────────────────────────────────────────────────────────
UPDATE_SHARE = r"C:\Users\ncurtis\Documents\PROJECTS\!Completed Programs\Study Aggregator\Study Aggregator 2.0\update-tester"  # TODO: set actual UNC path
AUMID = "Ronsin.StudyAggregator"
INSTALLER_NAME = "StudyAggregatorSetup.exe"
BLOCKING_PROCESSES = {"study aggregator.exe", "7z.exe", "7za.exe"}

_exit_event = threading.Event()


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
            creationflags=subprocess.CREATE_NO_WINDOW,
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


def on_activated(args):
    if getattr(args, "arguments", "") == "update":
        install_update()
    _exit_event.set()


def on_dismissed(args):
    _exit_event.set()


def on_failed(args):
    _exit_event.set()


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

    toast.on_activated = on_activated
    toast.on_dismissed = on_dismissed
    toast.on_failed = on_failed

    notifier = InteractableWindowsToaster("Study Aggregator", notifierAUMID=AUMID)
    notifier.show_toast(toast)

    # Keep alive until button clicked, dismissed, or 5 min safety timeout
    _exit_event.wait(timeout=300)


if __name__ == "__main__":
    main()
