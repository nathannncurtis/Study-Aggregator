import os
import pydicom
import clipboard
from collections import defaultdict
import pyzipper
import shutil
import tempfile
from PyQt5.QtWidgets import QApplication, QMessageBox, QInputDialog, QLineEdit, QProgressBar, QLabel, QVBoxLayout, QWidget
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon
import time
import sys
import ctypes
import zipfile
import concurrent.futures
import gc
from functools import lru_cache
import mmap
import logging
import traceback
import subprocess
import re

# --- Logging Setup ---
def setup_logging():
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    for handler in logger.handlers[:]:
        if isinstance(handler, logging.StreamHandler) and handler.stream in [sys.stdout, sys.stderr]:
            logger.removeHandler(handler)

    try:
        log_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dicom_aggregator.log")
        
        has_file_handler = any(isinstance(h, logging.FileHandler) and h.baseFilename == os.path.abspath(log_file_path) for h in logger.handlers)
        
        if not has_file_handler:
            file_handler = logging.FileHandler(log_file_path, mode='a')
            file_handler.setFormatter(log_formatter)
            logger.addHandler(file_handler)
        else:
            pass
            
    except Exception as e:
        print(f"CRITICAL: Error setting up file logger: {e}", file=sys.stderr)

    logging.info("Logging initialized. Output directed to dicom_aggregator.log")

# --- End Logging Setup ---

user32 = ctypes.WinDLL('user32', use_last_error=True)

def set_busy_cursor():
    try:
        user32.SetSystemCursor(user32.LoadCursorW(0, 32514), 32512)
    except Exception as e:
        logging.warning(f"Could not set busy cursor: {e}")

def reset_cursor():
    try:
        user32.SystemParametersInfoW(87, 0, None, 0)
    except Exception as e:
        logging.warning(f"Could not reset cursor: {e}")

# Determine icon_path safely
try:
    if getattr(sys, 'frozen', False):
        application_path = os.path.dirname(sys.executable)
    else:
        application_path = os.path.dirname(os.path.abspath(__file__))
    icon_path = os.path.join(application_path, 'agg.ico')
except Exception:
    icon_path = 'agg.ico'

# 7zip detection and path finding
def find_7zip():
    """Find 7zip executable on the system"""
    possible_paths = [
        r"C:\Program Files\7-Zip\7z.exe",
        r"C:\Program Files (x86)\7-Zip\7z.exe",
        "7z",  # In PATH
        "7za", # Standalone version
    ]
    
    for path in possible_paths:
        try:
            result = subprocess.run([path], capture_output=True, timeout=5)
            logging.info(f"Found 7zip at: {path}")
            return path
        except (subprocess.TimeoutExpired, subprocess.CalledProcessError, FileNotFoundError, OSError):
            continue
    
    logging.info("7zip not found on system, will use pyzipper fallback")
    return None

# Global 7zip path
SEVEN_ZIP_PATH = find_7zip()

class ProgressDialog(QWidget):
    def __init__(self, title="Processing"):
        super().__init__()
        self.setWindowTitle(title)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))
        else:
            logging.warning(f"Icon file not found at {icon_path}, not setting window icon for ProgressDialog.")
        
        layout = QVBoxLayout()
        
        self.label = QLabel("Processing files...")
        layout.addWidget(self.label)
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        layout.addWidget(self.progress_bar)
        
        self.setLayout(layout)
        self.resize(400, 100)
        
    def update_progress(self, value, text=None):
        self.progress_bar.setValue(value)
        if text:
            self.label.setText(text)
            
    def closeEvent(self, event):
        event.ignore()

def show_error_popup(message):
    logging.error(f"Displaying error popup: {message}")
    try:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        msg.setWindowTitle("Error")
        if os.path.exists(icon_path):
            msg.setWindowIcon(QIcon(icon_path))
        else:
            logging.warning(f"Icon file not found at {icon_path}, not setting window icon for error popup.")
        msg.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg.exec_()
    except Exception as e:
        logging.error(f"Failed to show error popup: {e}", exc_info=True)

def show_success_popup(message):
    logging.info(f"Displaying success popup: {message}")
    try:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(message)
        msg.setWindowTitle("Success")
        if os.path.exists(icon_path):
            msg.setWindowIcon(QIcon(icon_path))
        else:
            logging.warning(f"Icon file not found at {icon_path}, not setting window icon for success popup.")
        msg.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg.exec_()
    except Exception as e:
        logging.error(f"Failed to show success popup: {e}", exc_info=True)

def is_valid_dicom_file(file_path):
    try:
        if not os.path.exists(file_path) or os.path.isdir(file_path):
            return False
            
        # Skip PDF files and other common non-DICOM file types
        file_ext = os.path.splitext(file_path)[1].lower()
        if file_ext in ['.pdf', '.txt', '.exe', '.bat', '.inf', '.chm']:
            logging.debug(f"Skipping non-DICOM file type: {file_path}")
            return False
            
        if os.path.getsize(file_path) < 132:
            return False
            
        with open(file_path, 'rb') as f:
            with mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                if len(mm) < 132:
                    return False
                mm.seek(128)
                if mm.read(4) == b'DICM':
                    return True
    except IOError as e:
        logging.debug(f"IOError/mmap error during quick check for {file_path}: {e}. Falling back to pydicom.")
    except ValueError as e:
        logging.debug(f"ValueError (mmap) during quick check for {file_path}: {e}. Falling back to pydicom.")
    except Exception as e:
        logging.debug(f"Unexpected error during mmap check for {file_path}: {e}. Falling back to pydicom.")
        
    try:
        pydicom.dcmread(file_path, force=True, stop_before_pixels=True)
        return True
    except Exception as e:
        logging.debug(f"Pydicom could not read {file_path} as DICOM: {e}")
        return False

def normalize_name(name):
    """Normalize patient name for matching"""
    if not name or name == "Unknown":
        return None
    
    # Remove common separators and extra spaces
    name = str(name).replace("^", " ").replace(",", " ").replace("_", " ")
    name = " ".join(name.split())  # Remove extra whitespace
    
    # Split into parts and sort to handle "LAST FIRST" vs "FIRST LAST"
    parts = [part.strip().upper() for part in name.split() if part.strip()]
    if len(parts) >= 2:
        return tuple(sorted(parts))  # Return sorted tuple for matching
    return tuple(parts) if parts else None

def names_match(name1, name2):
    """Check if two names likely refer to the same person"""
    norm1 = normalize_name(name1)
    norm2 = normalize_name(name2)
    
    if norm1 is None or norm2 is None:
        return False
    
    return norm1 == norm2

@lru_cache(maxsize=2000)
def extract_study_info(dicom_file):
    try:
        if os.path.getsize(dicom_file) < 132:
            logging.debug(f"File {dicom_file} too small to be DICOM.")
            return None

        ds = pydicom.dcmread(dicom_file, force=True, specific_tags=[
            "StudyDate", "StudyDescription", "SeriesDescription", "Modality",
            "PatientName", "PatientBirthDate", "PatientID", "StudyInstanceUID", 
            "SeriesInstanceUID", "SeriesNumber"
        ], stop_before_pixels=True)

        patient_id = str(ds.get("PatientID", "")).strip()
        study_date = str(ds.get("StudyDate", "")).strip()
        study_description = str(ds.get("StudyDescription", "")).strip()
        series_description = str(ds.get("SeriesDescription", "")).strip()
        modality = str(ds.get("Modality", "")).strip()
        study_instance_uid = str(ds.get("StudyInstanceUID", "")).strip()
        series_instance_uid = str(ds.get("SeriesInstanceUID", "")).strip()
        series_number = str(ds.get("SeriesNumber", "")).strip()

        # Normalize patient name
        patient_name_raw = ds.get("PatientName", "")
        patient_name = str(patient_name_raw).replace("^", " ").strip()
        patient_name = " ".join(patient_name.split()) if patient_name else "Unknown"

        # Format DOB to MM-DD-YYYY
        patient_dob = str(ds.get("PatientBirthDate", "")).strip()
        if len(patient_dob) == 8:
            patient_dob = f"{patient_dob[4:6]}-{patient_dob[6:8]}-{patient_dob[0:4]}"
        else:
            patient_dob = "Unknown"

        # Format StudyDate to MM-DD-YYYY
        if len(study_date) == 8:
            study_date = f"{study_date[4:6]}-{study_date[6:8]}-{study_date[0:4]}"
        else:
            study_date = "Unknown"

        # Build base description (study level)
        base_description = study_description if study_description else "Study"
        
        # Build series description
        if series_description:
            series_desc = series_description
        else:
            series_desc = f"Series {series_number}" if series_number else "Unknown Series"

        # Return raw data for study/series grouping
        result = {
            'patient_id': patient_id or None,
            'patient_name': patient_name,
            'patient_dob': patient_dob,
            'study_date': study_date,
            'study_description': base_description,
            'series_description': series_desc,
            'modality': modality,
            'study_instance_uid': study_instance_uid or None,
            'series_instance_uid': series_instance_uid or None,
            'series_number': series_number or None
        }
        
        # Add more detailed logging for debugging
        if base_description == "Study" or patient_name == "Unknown":
            logging.warning(f"Potentially problematic DICOM file {dicom_file}: "
                          f"StudyDesc='{study_description}', PatientName='{patient_name_raw}', "
                          f"PatientID='{patient_id}', StudyUID='{study_instance_uid}'")
        else:
            logging.debug(f"Processed {os.path.basename(dicom_file)}: {patient_name} - {base_description}")
        
        return result

    except Exception as e:
        logging.warning(f"Failed to extract study info from {dicom_file}: {e}", exc_info=True)
        return None

def get_password_from_gui(parent_widget, description_text):
    password, ok = QInputDialog.getText(parent_widget, 'Password Required', 
                                        f'Enter password for {description_text}:', 
                                        QLineEdit.Password)
    if ok and password:
        return password.encode('utf-8')
    elif ok and not password:
        return b'' 
    return None

def process_dicom_file(file_path):
    try:
        if is_valid_dicom_file(file_path):
            return extract_study_info(file_path), file_path
    except Exception as e:
        logging.error(f"Unexpected error in process_dicom_file for {file_path}: {e}", exc_info=True)
    return None

def detect_zip_encryption_type(zip_path):
    """Detect if zip uses traditional or AES encryption"""
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            for zinfo in zf.infolist():
                if zinfo.flag_bits & 0x1:  # Encrypted
                    # Try to determine if it's AES by looking at extra field
                    if zinfo.extra:
                        # AES encryption typically has specific extra field signatures
                        if b'\x01\x99' in zinfo.extra:  # AES extra field signature
                            return "aes"
                    return "traditional"
            return "none"
    except Exception as e:
        logging.error(f"Error detecting encryption type for {zip_path}: {e}")
        return "unknown"

def extract_with_7zip(zip_path, extract_to, password=None):
    """Extract zip file using 7zip"""
    if not SEVEN_ZIP_PATH:
        return False
    
    try:
        cmd = [SEVEN_ZIP_PATH, 'x', zip_path, f'-o{extract_to}', '-y']
        if password:
            cmd.append(f'-p{password.decode("utf-8")}')
        
        logging.info(f"Extracting with 7zip: {' '.join(cmd[:-1])}{'[password]' if password else ''}")
        
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=300)
        
        if result.returncode == 0:
            logging.info(f"7zip extraction successful for {zip_path}")
            return True
        else:
            logging.warning(f"7zip extraction failed for {zip_path}: {result.stderr}")
            if "wrong password" in result.stderr.lower():
                raise Exception("WRONG_PASSWORD")
            return False
            
    except subprocess.TimeoutExpired:
        logging.error(f"7zip extraction timed out for {zip_path}")
        return False
    except Exception as e:
        logging.error(f"7zip extraction error for {zip_path}: {e}")
        if "WRONG_PASSWORD" in str(e):
            raise
        return False

def process_zip_file(zip_path, password=None, max_workers=None, progress_callback=None, nested_level=0, max_nested_level=5):
    found_studies = []
    if nested_level > max_nested_level:
        logging.warning(f"Maximum nested level ({max_nested_level}) reached for {zip_path}, skipping.")
        return found_studies
    
    try:
        temp_dir = tempfile.mkdtemp(prefix="dicom_agg_")
        logging.info(f"Extracting {zip_path} to temp dir: {temp_dir}")
    except Exception as e:
        logging.error(f"Failed to create temporary directory: {e}")
        if progress_callback:
            progress_callback(0, f"Error creating temp dir for {os.path.basename(zip_path)}")
        return found_studies

    if progress_callback:
        prefix = "  " * nested_level
        progress_callback(0, f"{prefix}Extracting {os.path.basename(zip_path)}...")
    
    extraction_successful = False
    encryption_type = detect_zip_encryption_type(zip_path)
    
    try:
        if encryption_type == "none":
            # No encryption, use standard zipfile
            logging.info(f"Extracting unencrypted zip: {zip_path}")
            with zipfile.ZipFile(zip_path, 'r') as zf:
                zf.extractall(temp_dir)
            extraction_successful = True
            
        elif encryption_type in ["traditional", "aes", "unknown"] and password is not None:
            # Try 7zip first for better performance
            if SEVEN_ZIP_PATH:
                try:
                    if extract_with_7zip(zip_path, temp_dir, password):
                        extraction_successful = True
                    else:
                        logging.info("7zip failed, falling back to pyzipper")
                except Exception as e:
                    if "WRONG_PASSWORD" in str(e):
                        raise Exception("WRONG_PASSWORD")
                    logging.info("7zip failed, falling back to pyzipper")
            
            # Fallback to pyzipper if 7zip failed or not available
            if not extraction_successful:
                logging.info(f"Using pyzipper for encrypted zip: {zip_path}")
                try:
                    if encryption_type == "traditional":
                        with zipfile.ZipFile(zip_path, 'r') as zf:
                            zf.extractall(temp_dir, pwd=password)
                    else:  # AES or unknown
                        with pyzipper.AESZipFile(zip_path) as zf:
                            zf.extractall(temp_dir, pwd=password)
                    extraction_successful = True
                except (RuntimeError, pyzipper.zipfile.BadZipFile) as e:
                    if "wrong password" in str(e).lower() or "WRONG_PASSWORD" in str(e).upper():
                        raise Exception("WRONG_PASSWORD")
                    logging.error(f"Pyzipper extraction failed: {e}")
                    
        else:
            # Encrypted but no password provided
            logging.warning(f"Zip file {zip_path} is encrypted, but no password was provided. Skipping extraction.")
            if progress_callback: 
                progress_callback(25, f"{prefix}Skipping encrypted {os.path.basename(zip_path)} (no password).")

    except Exception as e:
        if "WRONG_PASSWORD" in str(e):
            raise Exception("WRONG_PASSWORD")
        logging.error(f"Extraction failed for {zip_path}: {e}", exc_info=True)

    if not extraction_successful:
        shutil.rmtree(temp_dir, ignore_errors=True)
        logging.warning(f"Extraction failed for {zip_path}, cleaned up temp_dir.")
        return found_studies
            
    logging.info(f"Processing files from extracted {zip_path}...")
    if progress_callback:
        prefix = "  " * nested_level
        progress_callback(25, f"{prefix}Processing files from {os.path.basename(zip_path)}...")
    
    # Process nested zips
    nested_zip_files = []
    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.lower().endswith('.zip'):
                nested_zip_files.append(os.path.join(root, file))
    
    if nested_zip_files:
        logging.info(f"Found {len(nested_zip_files)} nested zip files in {zip_path}")
        nested_progress_start, nested_progress_end = 25, 65
        for i, nested_zip in enumerate(nested_zip_files):
            if progress_callback:
                prefix_nested_outer = "  " * nested_level
                nested_progress_val = nested_progress_start + (i / len(nested_zip_files) * (nested_progress_end - nested_progress_start))
                progress_callback(int(nested_progress_val), f"{prefix_nested_outer}Processing nested zip {i+1}/{len(nested_zip_files)}: {os.path.basename(nested_zip)}")
            def nested_progress_callback(percent, text):
                if progress_callback and percent >= 0:
                    range_size = (nested_progress_end - nested_progress_start) / len(nested_zip_files)
                    start_pos = nested_progress_start + (i * range_size)
                    scaled_percent = start_pos + (percent / 100) * range_size
                    prefix_nested_inner = "  " * (nested_level + 1)
                    progress_callback(int(scaled_percent), f"{prefix_nested_inner}{text}")
            nested_studies = process_zip_file(nested_zip, password, max_workers, nested_progress_callback, nested_level + 1, max_nested_level)
            found_studies.extend(nested_studies)
    
    # Process DICOM files
    progress_start_dicom = 65 if nested_zip_files else 25
    dicom_files = []
    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.lower().endswith('.zip'): continue
            file_path = os.path.join(root, file)
            ext = os.path.splitext(file)[1].lower()
            if ext in ('.dcm', '.ima', '.dicom', '') or not ext :
                dicom_files.append(file_path)
    
    logging.info(f"Found {len(dicom_files)} potential DICOM files in extracted {zip_path}")
    if dicom_files:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            batch_size = min(100, len(dicom_files)) if len(dicom_files) > 0 else 1
            num_batches = (len(dicom_files) + batch_size -1) // batch_size
            for batch_idx in range(num_batches):
                if progress_callback:
                    progress_percent = progress_start_dicom + ((batch_idx / num_batches) * (100 - progress_start_dicom))
                    prefix_dicom = "  " * nested_level
                    progress_callback(int(progress_percent), f"{prefix_dicom}Processing {os.path.basename(zip_path)}: batch {batch_idx+1}/{num_batches}...")
                
                start_idx = batch_idx * batch_size
                end_idx = start_idx + batch_size
                batch = dicom_files[start_idx:end_idx]
                futures = {executor.submit(process_dicom_file, f): f for f in batch}
                for future in concurrent.futures.as_completed(futures):
                    result = future.result()
                    if result:
                        study_info, file_path = result
                        if study_info:
                            study_info['source_path'] = zip_path
                            found_studies.append(study_info)
    
    logging.info(f"Found {len([s for s in found_studies if s])} valid studies in {zip_path}")
    if progress_callback:
        progress_callback(100, f"{'  ' * nested_level}Completed processing {os.path.basename(zip_path)}")
    
    try:
        shutil.rmtree(temp_dir)
        logging.info(f"Successfully removed temp directory: {temp_dir}")
    except Exception as e:
        logging.error(f"Error removing temp directory {temp_dir}: {e}")
    gc.collect()
    return found_studies

def find_zip_files(directory):
    zip_files = []
    try:
        for root, dirs, files in os.walk(directory):
            dirs[:] = [d for d in dirs if not d.startswith('.')] 
            for file in files:
                if file.lower().endswith('.zip'):
                    zip_files.append(os.path.join(root, file))
    except Exception as e:
        logging.error(f"Error scanning directory {directory} for zip files: {e}")
    return zip_files

def is_patient_all_unknown(patient_data):
    """Check if a patient has all unknown/empty identifying information"""
    patient_id = patient_data.get('patient_id')
    patient_name = patient_data.get('patient_name', 'Unknown')
    patient_dob = patient_data.get('patient_dob', 'Unknown')
    
    # Consider empty strings as unknown too
    id_unknown = not patient_id or patient_id.strip() == ''
    name_unknown = not patient_name or patient_name.strip() == '' or patient_name.strip() == 'Unknown'
    dob_unknown = not patient_dob or patient_dob.strip() == '' or patient_dob.strip() == 'Unknown'
    
    return id_unknown and name_unknown and dob_unknown

def merge_patients(studies_list):
    """Merge studies from the same patient based on ID and name matching, group by study and series"""
    patients = {}
    
    logging.info(f"Starting merge_patients with {len(studies_list)} studies")
    
    for i, study in enumerate(studies_list):
        if not study or not isinstance(study, dict):
            logging.warning(f"Invalid study data at index {i}: {study}")
            continue
            
        patient_id = study.get('patient_id')
        patient_name = study.get('patient_name', 'Unknown')
        patient_dob = study.get('patient_dob', 'Unknown')
        study_desc = study.get('study_description', 'Unknown')
        
        # Log problematic entries with source file information
        if patient_name == 'Unknown' or study_desc in ['Study', 'Unknown']:
            source_file = study.get('source_path', 'Unknown source')
            logging.warning(f"Study {i} has missing data from {source_file} - Name: '{patient_name}', "
                          f"StudyDesc: '{study_desc}', ID: '{patient_id}', DOB: '{patient_dob}'")
        
        # Find matching patient
        matched_key = None
        
        # First, try to match by Patient ID
        if patient_id:
            for key in patients.keys():
                if patients[key].get('patient_id') == patient_id:
                    matched_key = key
                    break
        
        # If no ID match, try name matching
        if not matched_key:
            for key in patients.keys():
                if names_match(patient_name, patients[key].get('patient_name')):
                    # Check for DOB conflicts (both real but different)
                    existing_dob = patients[key].get('patient_dob', 'Unknown')
                    if (patient_dob != 'Unknown' and existing_dob != 'Unknown' and 
                        patient_dob != existing_dob):
                        # DOB conflict - don't merge, will create separate entry
                        continue
                    matched_key = key
                    break
        
        if matched_key:
            # Merge with existing patient
            existing = patients[matched_key]
            
            # Update DOB if current study has one and existing doesn't
            if patient_dob != 'Unknown' and existing.get('patient_dob') == 'Unknown':
                existing['patient_dob'] = patient_dob
        else:
            # Create new patient entry
            patient_key = f"{patient_name}_{patient_dob}_{patient_id or 'NO_ID'}"
            patients[patient_key] = {
                'patient_id': patient_id,
                'patient_name': patient_name,
                'patient_dob': patient_dob,
                'studies': {}
            }
            matched_key = patient_key
            logging.debug(f"Created new patient entry: {patient_key}")
        
        # Group by study first
        patient_data = patients[matched_key]
        study_uid = study.get('study_instance_uid') or f"{study.get('study_date', 'Unknown')}_{study.get('study_description', 'Unknown')}"
        
        if study_uid not in patient_data['studies']:
            patient_data['studies'][study_uid] = {
                'study_date': study.get('study_date', 'Unknown'),
                'study_description': study.get('study_description', 'Unknown'),
                'all_series': set()  # Track all series in this study regardless of modality
            }
        
        # Track all series within the study (regardless of modality)
        series_uid = study.get('series_instance_uid') or f"{study.get('series_number', 'Unknown')}_{study.get('series_description', 'Unknown')}"
        patient_data['studies'][study_uid]['all_series'].add(series_uid)
    
    # Filter out patients with all unknown identifying information
    patients_before_filter = len(patients)
    patients = {k: v for k, v in patients.items() if not is_patient_all_unknown(v)}
    patients_after_filter = len(patients)
    
    if patients_before_filter > patients_after_filter:
        filtered_count = patients_before_filter - patients_after_filter
        logging.info(f"Filtered out {filtered_count} patients with all unknown identifying information")
    
    # Log final patient summary
    for patient_key, patient_data in patients.items():
        study_count = len(patient_data['studies'])
        logging.info(f"Patient '{patient_data.get('patient_name')}' (ID: {patient_data.get('patient_id')}) "
                    f"has {study_count} studies")
    
    return patients

class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, str)
    finished_signal = pyqtSignal(object)
    error_signal = pyqtSignal(str)

    def __init__(self, input_path, max_workers):
        super().__init__()
        self.input_path = input_path
        self.max_workers = max_workers
        self.max_nested_level = 5

    def run(self):
        try:
            set_busy_cursor()
            all_studies = []
            
            if os.path.isfile(self.input_path) and self.input_path.lower().endswith('.zip'):
                self.progress_updated.emit(0, f"Processing zip file: {os.path.basename(self.input_path)}")
                encryption_type = detect_zip_encryption_type(self.input_path)
                thread_password = None
                
                if encryption_type != "none":
                    self.progress_updated.emit(-1, "Password")
                    password_attr_set = False
                    logging.debug("Waiting for password attribute for single zip...")
                    wait_start_time = time.time()
                    while not password_attr_set and (time.time() - wait_start_time < 600):
                        if hasattr(self, 'password_from_gui'):
                            thread_password = self.password_from_gui
                            del self.password_from_gui
                            password_attr_set = True
                            logging.debug("Password attribute received for single zip")
                        else:
                            time.sleep(0.1)
                    if not password_attr_set:
                        logging.error("Timeout or failure waiting for password for single zip.")
                        self.error_signal.emit("Failed to get password for encrypted ZIP.")
                        reset_cursor()
                        return
                
                try:
                    study_data = process_zip_file(self.input_path, thread_password, self.max_workers,
                                                  lambda p, t: self.progress_updated.emit(p, t),
                                                  0, self.max_nested_level)
                    all_studies.extend(study_data)
                except Exception as e:
                    if "WRONG_PASSWORD" in str(e):
                        self.error_signal.emit("Wrong password provided for encrypted ZIP file.")
                    else:
                        self.error_signal.emit(f"Failed to extract ZIP file: {str(e)}")
                    reset_cursor()
                    return
                    
            elif os.path.isdir(self.input_path):
                self.progress_updated.emit(0, f"Processing directory: {os.path.basename(self.input_path)}")
                all_studies = self.extract_directory_info()
            else:
                err_msg = f"Invalid path or file type: {self.input_path}"
                logging.error(err_msg)
                self.error_signal.emit(err_msg)
                reset_cursor()
                return
            
            try:
                logging.debug(f"Processing complete. Found {len(all_studies)} total studies.")
                # Filter out any None entries
                valid_studies = [s for s in all_studies if s is not None]
                merged_patients = merge_patients(valid_studies)
                logging.debug(f"Merged into {len(merged_patients)} unique patients.")
                self.finished_signal.emit(merged_patients)
                logging.debug("Finished signal emitted successfully.")
            except Exception as emit_err:
                logging.error(f"Failed to emit finished signal: {emit_err}", exc_info=True)
        except Exception as e:
            logging.error(f"Error in ProcessingThread: {e}", exc_info=True)
            self.error_signal.emit(f"An unexpected error occurred: {str(e)}")
        finally:
            reset_cursor()
            
    def extract_directory_info(self):
        all_studies = []
        self.progress_updated.emit(0, "Scanning for zip files...")
        zip_files = find_zip_files(self.input_path)
        thread_shared_password = None
        
        # Process ZIP files if found
        if zip_files:
            self.progress_updated.emit(5, f"Found {len(zip_files)} zip files. Checking for encryption...")
            password_needed_flags = {}
            any_zip_needs_password = False
            
            for i, zip_path_check in enumerate(zip_files):
                try:
                    encryption_type = detect_zip_encryption_type(zip_path_check)
                    is_needed = encryption_type != "none"
                    password_needed_flags[zip_path_check] = is_needed
                    if is_needed: any_zip_needs_password = True
                    self.progress_updated.emit(5 + int((i / len(zip_files)) * 10), f"Checking zip file {i+1}/{len(zip_files)}...")
                except Exception as e_check:
                    logging.warning(f"Could not determine password status for {zip_path_check}: {e_check}. Assuming needed.")
                    password_needed_flags[zip_path_check] = True
                    any_zip_needs_password = True
            
            if any_zip_needs_password:
                self.progress_updated.emit(-1, "Password")
                shared_password_attr_set = False
                logging.debug("Waiting for shared_password attribute for directory...")
                wait_start_time = time.time()
                while not shared_password_attr_set and (time.time() - wait_start_time < 600):
                    if hasattr(self, 'shared_password_from_gui'):
                        thread_shared_password = self.shared_password_from_gui
                        del self.shared_password_from_gui
                        shared_password_attr_set = True
                        logging.debug(f"Shared_password attribute received: {'Yes' if thread_shared_password is not None else 'No/Cancelled'}")
                    else:
                        time.sleep(0.1)
                if not shared_password_attr_set:
                    logging.error("Timeout or failure waiting for shared password.")
            
            # Process ZIP files (use 60% of progress for ZIPs)
            for i, zip_path_process in enumerate(zip_files):
                current_zip_password_to_use = thread_shared_password if password_needed_flags.get(zip_path_process, False) else None
                progress_start_zip = 15 + int((i / len(zip_files)) * 45)  # 15-60%
                progress_end_zip = 15 + int(((i + 1) / len(zip_files)) * 45)
                def zip_progress_callback(percent, text):
                    if percent >= 0:
                        scaled_percent = progress_start_zip + int((percent / 100) * (progress_end_zip - progress_start_zip))
                        self.progress_updated.emit(scaled_percent, text)
                try:
                    self.progress_updated.emit(progress_start_zip, f"Processing zip {i+1}/{len(zip_files)}: {os.path.basename(zip_path_process)}")
                    found_studies_in_zip = process_zip_file(zip_path_process, current_zip_password_to_use, self.max_workers, 
                                                            zip_progress_callback, 0, self.max_nested_level)
                    all_studies.extend(found_studies_in_zip)
                except Exception as e_proc:
                    if "WRONG_PASSWORD" in str(e_proc):
                        logging.error(f"Wrong password for zip {zip_path_process}")
                        self.progress_updated.emit(progress_end_zip, f"Wrong password for {os.path.basename(zip_path_process)}")
                    else:
                        logging.error(f"Error processing zip {zip_path_process} in directory: {e_proc}", exc_info=True)
                        self.progress_updated.emit(progress_end_zip, f"Error with {os.path.basename(zip_path_process)}: {str(e_proc)[:50]}")
        
        # ALWAYS process loose DICOM files in addition to ZIPs (use remaining 40% of progress)
        progress_start_dicom = 60 if zip_files else 0
        self.progress_updated.emit(progress_start_dicom, "Scanning for loose DICOM files...")
        
        all_potential_dicom_files = []
        self.progress_updated.emit(progress_start_dicom + 5, "Analyzing file types for loose DICOMs...")
        file_scan_count, total_files_to_scan = 0, 0
        
        # Count total files first
        for root, _, files in os.walk(self.input_path):
            for file_name in files:
                if not file_name.lower().endswith('.zip'):  # Skip ZIP files we already processed
                    total_files_to_scan += 1
        
        # Find potential DICOM files
        for root, _, files in os.walk(self.input_path):
            for file_name in files:
                if file_name.lower().endswith('.zip'):  # Skip ZIP files
                    continue
                    
                file_scan_count += 1
                if file_scan_count % 200 == 0 and total_files_to_scan > 0:
                    progress_val = progress_start_dicom + 5 + int((file_scan_count / total_files_to_scan) * 15)
                    self.progress_updated.emit(progress_val, f"Analyzing file types: {file_scan_count}/{total_files_to_scan}")
                
                file_path, ext = os.path.join(root, file_name), os.path.splitext(file_name)[1].lower()
                if ext in ('.dcm', '.ima', '.dicom', '') or not ext:
                    all_potential_dicom_files.append(file_path)
        
        logging.info(f"Found {len(all_potential_dicom_files)} potential loose DICOM files")
        
        # Validate and process loose DICOM files
        dicom_files_confirmed = []
        if all_potential_dicom_files:
            self.progress_updated.emit(progress_start_dicom + 20, f"Validating {len(all_potential_dicom_files)} potential DICOM files...")
            for i, pf_path in enumerate(all_potential_dicom_files):
                if (i+1) % 50 == 0:
                    progress_val = progress_start_dicom + 20 + int(((i+1) / len(all_potential_dicom_files)) * 15)
                    self.progress_updated.emit(progress_val, f"Validating DICOMs: {i+1}/{len(all_potential_dicom_files)}")
                if is_valid_dicom_file(pf_path):
                    dicom_files_confirmed.append(pf_path)
        
        logging.info(f"Found {len(dicom_files_confirmed)} valid loose DICOM files")
        
        if dicom_files_confirmed:
            self.progress_updated.emit(progress_start_dicom + 35, f"Processing {len(dicom_files_confirmed)} loose DICOM files...")
            processed_count = 0
            with concurrent.futures.ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                futures = {executor.submit(process_dicom_file, f): f for f in dicom_files_confirmed}
                total_to_process = len(dicom_files_confirmed)
                for future in concurrent.futures.as_completed(futures):
                    processed_count += 1
                    progress_val = progress_start_dicom + 35 + int((processed_count / total_to_process) * 25)
                    self.progress_updated.emit(progress_val, f"Processing loose DICOMs: {processed_count}/{total_to_process}")
                    result = future.result()
                    if result:
                        study_info, file_path_processed = result
                        if study_info: 
                            study_info['source_path'] = file_path_processed
                            all_studies.append(study_info)
        
        self.progress_updated.emit(95, "Finalizing results from directory...")
        logging.info(f"Total studies found: {len(all_studies)} (from ZIPs and loose files)")
        return all_studies

def main_app_logic():
    cpu_cores = os.cpu_count() if os.cpu_count() is not None else 1
    max_workers = min(cpu_cores, 8) 
    logging.info(f"Using up to {max_workers} worker threads.")
    logging.info(f"7zip available: {'Yes' if SEVEN_ZIP_PATH else 'No'}")
    
    app = QApplication.instance() or QApplication(sys.argv)
    
    if len(sys.argv) < 2:
        logging.error("Usage: script.py <path_to_directory_or_zip_file>")
        print("Usage: script.py <path_to_directory_or_zip_file>", file=sys.stderr)
        return 1
        
    input_path = sys.argv[1]
    if not os.path.exists(input_path):
        logging.error(f"Input path does not exist: {input_path}")
        print(f"Error: Input path does not exist: {input_path}", file=sys.stderr)
        return 1
        
    progress_dialog = ProgressDialog("Processing DICOM Files")
    progress_dialog.show()
    
    processing_thread = ProcessingThread(input_path, max_workers)
    
    def update_progress_slot(value, text):
        if value == -1 and text == "Password":
            progress_dialog.hide()
            try:
                if os.path.isfile(processing_thread.input_path) and processing_thread.input_path.lower().endswith('.zip'):
                    logging.info(f"Requesting password for single zip: {processing_thread.input_path}")
                    password_bytes = get_password_from_gui(progress_dialog, f"encrypted ZIP: {os.path.basename(processing_thread.input_path)}")
                    processing_thread.password_from_gui = password_bytes
                else:
                    logging.info("Requesting shared password for directory processing.")
                    password_bytes = get_password_from_gui(progress_dialog, "any password-protected ZIP files in the directory")
                    processing_thread.shared_password_from_gui = password_bytes
            finally:
                progress_dialog.show()
        else:
            progress_dialog.update_progress(value, text)
    
    processing_thread.progress_updated.connect(update_progress_slot)
    
    def on_finished_slot(merged_patients):
        logging.debug("on_finished_slot() called.")
        try:
            progress_dialog.update_progress(100, "Processing complete. Formatting results...")
            logging.debug(f"Received merged_patients with {len(merged_patients)} entries")
            logging.debug(f"Merged patients type: {type(merged_patients)}")

            if not merged_patients:
                logging.warning("No patient data found after processing and filtering.")
                show_error_popup("Unable to aggregate DICOM studies: All extracted data contains insufficient patient identification information. Please manually inspect the source files to verify they contain valid DICOM headers with patient details.")
                progress_dialog.close()
                app.quit()
                return

            logging.debug("Patient data found. Beginning formatting.")

            # Sort patients by ID (treat all as strings for consistency), then by name
            def sort_patients(patient_data):
                pid = patient_data.get('patient_id') or 'ZZZZ'
                name = patient_data.get('patient_name', 'Unknown')
                
                # Convert all patient IDs to strings and pad numeric ones for proper sorting
                try:
                    if pid and pid.isdigit():
                        # Pad numeric IDs with leading zeros for proper string sorting
                        pid_val = pid.zfill(10)  # Pad to 10 digits
                    else:
                        pid_val = str(pid) if pid else 'ZZZZ'
                except:
                    pid_val = str(pid) if pid else 'ZZZZ'
                    
                return (pid_val, name.lower())

            sorted_patients = sorted(merged_patients.values(), key=sort_patients)
            
            lines = []
            for patient in sorted_patients:
                pid = patient.get('patient_id', '')
                name = patient.get('patient_name', 'Unknown')
                dob = patient.get('patient_dob', 'Unknown')
                
                # Format patient header
                if pid:
                    display_name = f"NAME: {name} DOB: {dob}, ID: {pid}"
                else:
                    display_name = f"NAME: {name} DOB: {dob}, ID: Unknown"

                lines.extend([f"{display_name}\r\n", "STUDIES\r\n\r\n"])

                # Get studies dictionary
                studies_dict = patient.get('studies', {})
                logging.debug(f"Studies type: {type(studies_dict)}, count: {len(studies_dict)}")
                
                if isinstance(studies_dict, dict) and studies_dict:
                    # Sort studies by date, then description
                    sorted_studies = sorted(studies_dict.values(), 
                                        key=lambda x: (x.get('study_date', 'Unknown'), x.get('study_description', '')))
                    
                    for study in sorted_studies:
                        study_date = study.get('study_date', 'Unknown')
                        study_desc = study.get('study_description', 'Unknown')
                        
                        # Get all series in this study (regardless of modality)
                        all_series = study.get('all_series', set())
                        series_count = len(all_series)
                        
                        if series_count > 0:
                            # Format the line as: "DATE STUDY_DESCRIPTION (X series)"
                            line = f"{study_date} {study_desc} ({series_count} series)\r\n"
                            lines.append(line)
                        else:
                            # Fallback if no series data
                            lines.append(f"{study_date} {study_desc}\r\n")

                lines.append("\r\n" + "="*50 + "\r\n\r\n")

            if not lines:
                show_error_popup("Unable to aggregate DICOM studies: No valid patient data could be formatted for output. Please manually inspect the source files.")
                progress_dialog.close()
                app.quit()
                return

            # Copy to clipboard
            output_text = "".join(lines)
            logging.debug(f"Final output preview: {output_text[:500]}...")
            
            try:
                clipboard.copy(output_text)
                progress_dialog.close()
                show_success_popup("Studies have been successfully copied to the clipboard")
                app.quit()
            except Exception as clipboard_error:
                progress_dialog.close()
                show_error_popup(f"Failed to copy results to clipboard: {str(clipboard_error)}")
                app.quit()
                
        except Exception as e_format:
            logging.error(f"Error in on_finished_slot: {e_format}", exc_info=True)
            progress_dialog.close()
            show_error_popup(f"Error formatting results: {str(e_format)}")
            app.quit()

    def on_error_slot(error_message):
        logging.error(f"Processing error signal received: {error_message}")
        progress_dialog.close()
        show_error_popup(error_message)
        app.quit()
    
    processing_thread.finished_signal.connect(on_finished_slot)
    processing_thread.error_signal.connect(on_error_slot)
    
    logging.info(f"Starting processing thread for input: {input_path}")
    processing_thread.start()
    return app.exec_()

def profile_main():
    start_time = time.time()
    setup_logging()
    logging.info("Python garbage collection is currently disabled for the main application logic.")
    logging.warning("Garbage collection is disabled. This may increase peak memory usage but can speed up processing for this utility.")
    exit_code = main_app_logic()
    total_duration = time.time() - start_time
    logging.info(f"Total execution time: {total_duration:.2f} seconds")
    return exit_code

if __name__ == "__main__":
    if sys.platform.startswith('win'):
        import multiprocessing
        multiprocessing.freeze_support()
    gc.disable()
    final_exit_code = profile_main()
    gc.enable()
    logging.info("Python garbage collection re-enabled.")
    logging.info("Application finished.")
    sys.exit(final_exit_code)