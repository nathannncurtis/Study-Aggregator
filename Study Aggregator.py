import os
import pydicom
import clipboard
from datetime import datetime
from collections import defaultdict
import pyzipper
import shutil
import tempfile
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QLineEdit
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import time
import sys
import ctypes
import zipfile

user32 = ctypes.WinDLL('user32', use_last_error=True)

def set_busy_cursor():
    user32.SetSystemCursor(user32.LoadCursorW(0, 32514), 32512)

def reset_cursor():
    user32.SystemParametersInfoW(87, 0, None, 0)

icon_path = os.path.join(os.path.dirname(__file__), 'agg.ico')

def show_error_popup(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText(message)
    msg.setWindowTitle("Error")
    msg.setWindowIcon(QIcon(icon_path))
    msg.setWindowFlag(Qt.WindowStaysOnTopHint)
    msg.exec_()

def show_success_popup(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText(message)
    msg.setWindowTitle("Success")
    msg.setWindowIcon(QIcon(icon_path))
    msg.setWindowFlag(Qt.WindowStaysOnTopHint)
    msg.exec_()

def is_valid_dicom_file(file_path):
    with open(file_path, 'rb') as f:
        f.seek(128)
        return f.read(4) == b'DICM'

def extract_study_info(dicom_file):
    ds = pydicom.dcmread(dicom_file, force=True)
    
    if not hasattr(ds, 'PixelData'):
        return None

    study_date = ds.get("StudyDate", "Unknown")
    study_description = ds.get("StudyDescription", "Unknown")
    series_description = ds.get("SeriesDescription", "Unknown")

    if study_date != "Unknown":
        study_date = datetime.strptime(study_date, "%Y%m%d").strftime("%m-%d-%Y")

    if study_description == "Unknown" or not study_description.strip():
        for element in ds:
            if element.VR == "LO" or element.VR == "SH":
                if "Study" in element.name or "Series" in element.name:
                    study_description = element.value
                    break

    if (study_description == "Unknown" or not study_description.strip()) and series_description != "Unknown":
        study_description = series_description

    if study_description and len(study_description) > 10 and study_description.isalnum():
        study_description = "Unknown"

    if not study_description.strip():
        study_description = "Unknown"

    return (study_date, study_description)

def get_password_from_gui(zip_path):
    password, ok = QInputDialog.getText(None, 'Password Required', f'Enter password for encrypted zip file {zip_path}:', QLineEdit.Password)
    if ok:
        return password.encode()  # Return the password encoded as bytes
    return None

def process_zip_file(zip_path, app):
    found_studies = defaultdict(set)  # Change to defaultdict for correct structure
    password = get_password_from_gui(zip_path)

    if password is None:
        print("No password provided.")
        return found_studies  # Return an empty defaultdict if no password is provided

    temp_dir = tempfile.mkdtemp()
    print(f"Extracting {zip_path} to temporary directory...")

    try:
        with pyzipper.AESZipFile(zip_path) as zf:
            zf.extractall(temp_dir, pwd=password)
    except (RuntimeError, zipfile.BadZipFile):
        print("Failed to extract zip file, incorrect password or corruption.")
        shutil.rmtree(temp_dir)
        return found_studies

    print(f"Processing extracted files from {zip_path}...")
    for root, _, files in os.walk(temp_dir):
        for file in files:
            file_path = os.path.join(root, file)
            app.processEvents()

            if file.endswith(".dcm") or not os.path.splitext(file)[1]:
                try:
                    study_info = extract_study_info(file_path)
                    if study_info:
                        found_studies[study_info].add(zip_path)  # Add zip path to study info
                except:
                    continue

    shutil.rmtree(temp_dir)
    return found_studies  # Return defaultdict (which behaves like a dict)

def extract_directory_info(directory, app):
    study_data = defaultdict(set)
    zip_files = []

    # 1. Find all zip files first
    print("Scanning for zip files...")
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith('.zip'):
                zip_files.append(os.path.join(root, file))

    # 2. Process each zip file
    for zip_path in zip_files:
        print(f"\nProcessing zip file: {zip_path}")
        found_studies = process_zip_file(zip_path, app)
        for study_info, path in found_studies:
            study_data[study_info].add(path)

    # Only check for regular DICOM files if no zip files found
    if not zip_files:
        print("No zip files found, scanning for individual DICOM files...")
        for root, _, files in os.walk(directory):
            for file in files:
                file_path = os.path.join(root, file)
                app.processEvents()
                
                if file.endswith(".dcm") or is_valid_dicom_file(file_path):
                    try:
                        study_info = extract_study_info(file_path)
                        if study_info:
                            study_data[study_info].add(file_path)
                    except:
                        continue

    # Only show error if no studies found after all processing
    if not study_data:
        show_error_popup("No valid studies were found in the selected directory.")
        return False

    return study_data

def main():
    app = QApplication([])  # Ensure QApplication is running
    set_busy_cursor()

    if len(sys.argv) < 2:
        print("Usage: script.py <path_to_directory_or_zip_file>")
        reset_cursor()
        sys.exit(1)

    input_path = sys.argv[1]

    try:
        start_time = time.time()

        # Check if the input path is a zip file or a directory
        if os.path.isfile(input_path) and input_path.endswith('.zip'):
            print(f"Processing single zip file: {input_path}")
            study_data = process_zip_file(input_path, app)
        elif os.path.isdir(input_path):
            print(f"Processing directory: {input_path}")
            study_data = extract_directory_info(input_path, app)
        else:
            print(f"Invalid path: {input_path}")
            reset_cursor()
            sys.exit(1)

        elapsed_time = time.time() - start_time
        print(f"Processing completed in {elapsed_time:.2f} seconds.")

        if not study_data:
            return

        # Sort and prepare study data for clipboard
        sorted_study_data = sorted(study_data.items(), key=lambda x: (x[0][0], x[0][1]))
        output_text = "The below studies are available\r\n\r\n"
        for (study_date, study_description), paths in sorted_study_data:
            if study_description != "Unknown" and not study_description.startswith("Study for"):
                output_text += f"{study_date} {study_description} \r\n"

        # Copy to clipboard
        clipboard.copy(output_text)
        print("Studies copied")

        show_success_popup("Studies have been successfully copied to the clipboard.")
    finally:
        reset_cursor()
        app.processEvents()

    app.quit()

if __name__ == "__main__":
    main()
