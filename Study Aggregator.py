import os
import pydicom
import clipboard
from datetime import datetime
from collections import defaultdict
import pyzipper
import shutil
import tempfile
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QIcon
import time
import sys  
import ctypes  

# Load Windows user32.dll for cursor manipulation
user32 = ctypes.WinDLL('user32', use_last_error=True)

# Function to set the system-wide cursor to busy (hourglass/spinning)
def set_busy_cursor():
    user32.SetSystemCursor(user32.LoadCursorW(0, 32514), 32512)  # 32514 is IDC_WAIT (busy cursor)

# Function to reset the system-wide cursor back to the default
def reset_cursor():
    user32.SystemParametersInfoW(87, 0, None, 0)  # 87 is SPI_SETCURSORS, which resets all cursors to default

# Relative path to ico file
icon_path = os.path.join(os.path.dirname(__file__), 'agg.ico')

# Function to display an error popup using PyQt5, always on top
def show_error_popup(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Critical)
    msg.setText(message)
    msg.setWindowTitle("Error")
    msg.setWindowIcon(QIcon(icon_path))
    msg.setWindowFlag(Qt.WindowStaysOnTopHint)
    msg.exec_()

# Function to display a success popup using PyQt5, always on top
def show_success_popup(message):
    msg = QMessageBox()
    msg.setIcon(QMessageBox.Information)
    msg.setText(message)
    msg.setWindowTitle("Success")
    msg.setWindowIcon(QIcon(icon_path))
    msg.setWindowFlag(Qt.WindowStaysOnTopHint)
    msg.exec_()

# Function to check if a file is a DICOM file
def is_valid_dicom_file(file_path):
    try:
        with open(file_path, 'rb') as f:
            f.seek(128)
            return f.read(4) == b'DICM'
    except Exception:
        return False

# Function to extract and format study information from a DICOM file
def extract_study_info(dicom_file):
    try:
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
    except pydicom.errors.InvalidDicomError as e:
        print(f"Invalid DICOM file {dicom_file}: {e}")
    except FileNotFoundError as e:
        print(f"File not found: {dicom_file}. Error: {e}")
    except Exception as e:
        print(f"Error reading {dicom_file}: {e}")
    return None

# Function to process DICOM files from zip archives
def process_zip_file(zip_path, study_data, app):
    try:
        with pyzipper.AESZipFile(zip_path, 'r') as zip_ref:
            password = None
            is_encrypted = False

            for file_info in zip_ref.infolist():
                if file_info.flag_bits & 0x1:
                    is_encrypted = True
                    print(f"File {file_info.filename} is encrypted.")
                    break

            if is_encrypted:
                password = input(f"Enter password for encrypted zip file {zip_path}: ")

            temp_dir = tempfile.mkdtemp()

            try:
                print(f"Extracting files from {zip_path} to {temp_dir}...")

                for file_name in zip_ref.namelist():
                    try:
                        extract_path = os.path.join(temp_dir, file_name)

                        if file_name.endswith('/'):
                            os.makedirs(extract_path, exist_ok=True)
                            continue

                        os.makedirs(os.path.dirname(extract_path), exist_ok=True)

                        if is_encrypted and password:
                            with zip_ref.open(file_name, pwd=password.encode()) as source_file:
                                with open(extract_path, 'wb') as target_file:
                                    shutil.copyfileobj(source_file, target_file)
                        else:
                            with zip_ref.open(file_name) as source_file:
                                with open(extract_path, 'wb') as target_file:
                                    shutil.copyfileobj(source_file, target_file)

                    except Exception as e:
                        print(f"Error extracting {file_name}: {e}")

                print(f"Processing extracted files in {temp_dir}...")

                for root, _, files in os.walk(temp_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        if file.endswith(".dcm") or not os.path.splitext(file)[1]:
                            study_info = extract_study_info(file_path)
                            if study_info:
                                study_data[study_info].add(zip_path)

            finally:
                print(f"Cleaning up temporary directory {temp_dir}...")
                shutil.rmtree(temp_dir)

    except FileNotFoundError as e:
        print(f"File not found: {zip_path}. Error: {e}")
    except Exception as e:
        print(f"Error processing zip file {zip_path}: {e}")

# Iterate through the directory and extract information from each DICOM file
def extract_directory_info(directory, app):
    study_data = defaultdict(set)

    for root, _, files in os.walk(directory):
        for file in files:
            file_path = os.path.join(root, file)
            app.processEvents()  # Keep UI responsive
            
            if file.lower().endswith(('.dcm', '.zip')) or is_valid_dicom_file(file_path):
                if file.endswith(".dcm") or not os.path.splitext(file)[1]:
                    study_info = extract_study_info(file_path)
                    if study_info:
                        study_data[study_info].add(file_path)
                elif file.endswith(".zip"):
                    process_zip_file(file_path, study_data, app)

    if not study_data:
        show_error_popup("No valid studies were found in the selected directory.")
        return False

    return study_data

# Main function to process DICOM files
def main():
    app = QApplication([])

    # Set the global busy cursor
    set_busy_cursor()

    # Check if the path is provided as a command-line argument
    if len(sys.argv) < 2:
        print("Usage: script.py <path_to_directory>")
        reset_cursor()  # Reset the cursor if there's an error
        sys.exit(1)

    dicom_directory = sys.argv[1]  # Get the directory path from command-line arguments

    try:
        start_time = time.time()

        study_data = extract_directory_info(dicom_directory, app)

        elapsed_time = time.time() - start_time
        print(f"Processing completed in {elapsed_time:.2f} seconds.")

        if not study_data:
            return

        sorted_study_data = sorted(study_data.items(), key=lambda x: (x[0][0], x[0][1]))
        output_text = "The below studies are available\r\n\r\n"
        for (study_date, study_description), paths in sorted_study_data:
            if study_description != "Unknown" and not study_description.startswith("Study for"):
                output_text += f"{study_date} {study_description} \r\n"

        clipboard.copy(output_text)
        print("Studies copied")

        show_success_popup("Studies have been successfully copied to the clipboard.")
    finally:
        reset_cursor()  # Reset the cursor back to default after processing
        app.processEvents()  # Ensure the cursor restoration takes effect

    # Exit after the success popup is acknowledged
    app.quit()

if __name__ == "__main__":
    main()
