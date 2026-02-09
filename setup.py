from cx_Freeze import setup, Executable
import sys
import os

with open("version.txt", "r") as _vf:
    APP_VERSION = _vf.read().strip()

# Force inclusion of reportlab
build_exe_options = {
    "packages": [
        # Core Python modules
        "os", "sys", "time", "gc", "threading", "signal", "subprocess", "re",
        "tempfile", "shutil", "io", "mmap", "traceback", "logging",

        # Collections and data structures
        "collections", "functools", "queue",

        # Concurrency
        "concurrent.futures", "multiprocessing",

        # Archive handling
        "zipfile", "pyzipper",

        # DICOM processing - include necessary modules
        "pydicom", "pydicom.filebase", "pydicom.filereader", "pydicom.dataset",
        "pydicom.dataelem", "pydicom.tag", "pydicom.uid", "pydicom.valuerep",
        "pydicom.multival", "pydicom.values", "pydicom.charset", "pydicom.config",
        "pydicom.datadict", "pydicom.errors", "pydicom.encaps",
        # Include pixel handlers to avoid config import errors
        "pydicom.pixel_data_handlers", "pydicom.pixel_data_handlers.numpy_handler",
        "pydicom.pixel_data_handlers.pillow_handler", "pydicom.pixel_data_handlers.jpeg_ls_handler",
        "pydicom.pixel_data_handlers.gdcm_handler", "pydicom.pixel_data_handlers.pylibjpeg_handler",

        # NumPy - required by pydicom
        "numpy",

        # GUI framework
        "PyQt5.QtWidgets", "PyQt5.QtCore", "PyQt5.QtGui", "PyQt5.QtPrintSupport",

        # PDF manipulation
        "pypdf",

        # Utilities
        "clipboard", "ctypes",

        # Toast notifications
        "windows_toasts",

        # Windows-specific
        "ctypes.wintypes" if sys.platform == "win32" else None,
    ],
    "excludes": [
        # Exclude unnecessary modules to reduce size
        "tkinter", "matplotlib", "pandas", "scipy",
        "PIL", "opencv", "tensorflow", "torch",
        # Only exclude the most problematic modules
        "pydicom.examples", "pydicom.data.data_manager", "pydicom.data.download"
    ],
    "include_files": [
        "agg.ico",
        "bd.pdf",
        "version.txt"
    ],
    "optimize": 2,
    "zip_include_packages": ["*"],
    "zip_exclude_packages": []
}

# Remove None entries from packages list
build_exe_options["packages"] = [pkg for pkg in build_exe_options["packages"] if pkg is not None]

# Suppress terminal window for GUI executables
# cx_Freeze 7+ renamed "Win32GUI" to "gui"
import cx_Freeze
_cx_major = int(cx_Freeze.__version__.split(".")[0]) if hasattr(cx_Freeze, "__version__") else 6
base_gui = ("gui" if _cx_major >= 7 else "Win32GUI") if sys.platform == "win32" else None

executables = [
    Executable("Study Aggregator.py", base=base_gui, icon="agg.ico"),
    Executable("reg.py", base=None, icon="agg.ico"),
    Executable("unreg.py", base=None, icon="agg.ico"),
    Executable("update_checker.py", base=base_gui, icon="agg.ico"),
]

setup(
    name="Study Aggregator",
    version=APP_VERSION,
    description="Study Aggregator for DICOM and Zip Directories with PDF Report Generation and Dark Theme UI",
    options={"build_exe": build_exe_options},
    executables=executables
)