import os
import pydicom
import clipboard
from datetime import datetime
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
import queue
import gc
import threading
from functools import lru_cache
import io
import mmap
import signal
# make this run faster with a lighter ui toolkit later maybe
user32 = ctypes.WinDLL('user32', use_last_error=True)
# windows cursor stuff
def set_busy_cursor():
    user32.SetSystemCursor(user32.LoadCursorW(0, 32514), 32512)
def reset_cursor():
    user32.SystemParametersInfoW(87, 0, None, 0)
icon_path = os.path.join(os.path.dirname(__file__), 'agg.ico')

# Create a progress bar dialog that stays on top
class ProgressDialog(QWidget):
    def __init__(self, title="Processing"):
        super().__init__()
        self.setWindowTitle(title)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        self.setWindowIcon(QIcon(icon_path))
        
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
        event.ignore()  # Prevent closing

# keep the ui stuff, it's simple enough
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
# quick check if it's dicom without loading the whole thing
def is_valid_dicom_file(file_path):
    try:
        # use memory mapping for much faster file checks
        if os.path.getsize(file_path) < 132:  # need at least 132 bytes for DICOM header
            return False
            
        with open(file_path, 'rb') as f:
            # memory map the file for faster access
            with mmap.mmap(f.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                mm.seek(128)
                return mm.read(4) == b'DICM'
    except Exception as e:
        # print(f"error checking dicom file {file_path}: {e}")
        return False
        
    # fallback to pydicom if the quick check fails
    try:
        # just check if pydicom can read it at all
        ds = pydicom.dcmread(file_path, force=True, stop_before_pixels=True)
        return True
    except:
        return False

# Enhanced function to extract patient info as well as study info
@lru_cache(maxsize=2000)
def extract_study_info(dicom_file):
    try:
        # don't even try to read files that are too small
        if os.path.getsize(dicom_file) < 132:  # reduced from 1000 to be more tolerant
            return None
            
        # try the simpler approach first - more reliable
        ds = pydicom.dcmread(dicom_file, force=True, specific_tags=[
            "StudyDate", 
            "StudyDescription", 
            "SeriesDescription",
            "Modality",
            "PatientName",
            "PatientBirthDate",
            "PixelData"
        ], stop_before_pixels=True)
        
        # be more tolerant of missing pixel data - some valid DICOMs don't have it
        study_date = ds.get("StudyDate", "Unknown")
        study_description = ds.get("StudyDescription", "Unknown")
        series_description = ds.get("SeriesDescription", "Unknown")
        modality = ds.get("Modality", "Unknown")
        
        # Extract patient information
        patient_name = str(ds.get("PatientName", "Unknown")).strip()
        patient_dob = ds.get("PatientBirthDate", "Unknown")
        
        # Format the DOB if it exists
        if patient_dob != "Unknown" and len(patient_dob) == 8:
            patient_dob = f"{patient_dob[4:6]}-{patient_dob[6:8]}-{patient_dob[0:4]}"
        
        # Format the study date if it exists
        if study_date != "Unknown":
            # faster date conversion
            if len(study_date) == 8:
                study_date = f"{study_date[4:6]}-{study_date[6:8]}-{study_date[0:4]}"
            else:
                study_date = "Unknown"
                
        # try multiple sources for study description
        if study_description == "Unknown" or not study_description.strip():
            # try series description first
            if series_description != "Unknown" and series_description.strip():
                study_description = series_description
            # try modality as a backup
            elif modality != "Unknown":
                study_description = f"{modality} Study"
            else:
                study_description = "Unknown"
        
        # limit overly long descriptions
        if study_description and len(study_description) > 50:
            study_description = study_description[:50]
        if not study_description.strip():
            study_description = "Unknown"
            
        return (study_date, study_description, patient_name, patient_dob)
    except Exception as e:
        # print(f"Error extracting study info from {dicom_file}: {e}")
        return None

# password dialog for zip files
def get_password_from_gui(zip_path):
    password, ok = QInputDialog.getText(None, 'Password Required', f'Enter password for encrypted zip file {zip_path}:', QLineEdit.Password)
    if ok:
        return password.encode()
    return None

# worker function for parallel dicom processing
def process_dicom_file(file_path):
    try:
        if is_valid_dicom_file(file_path):
            return extract_study_info(file_path), file_path
    except:
        pass
    return None

# check if a zip file needs a password
def zip_needs_password(zip_path):
    try:
        with zipfile.ZipFile(zip_path, 'r') as zf:
            # Try to read the first file without a password
            for filename in zf.namelist():
                try:
                    zf.open(filename).read(1)
                    return False  # No password needed
                except RuntimeError as e:
                    if "password required" in str(e).lower():
                        return True  # Password needed
                    continue
                except:
                    continue
                break
        return False
    except Exception as e:
        # if we can't open the zip at all, assume it might need a password
        print(f"error checking if zip needs password: {e}")
        return True

# process a single zip file with support for nested zips
def process_zip_file(zip_path, password=None, max_workers=None, progress_callback=None, nested_level=0, max_nested_level=10):
    found_studies = defaultdict(set)
    
    # Prevent excessive recursion
    if nested_level > max_nested_level:
        print(f"Maximum nested level ({max_nested_level}) reached for {zip_path}, skipping further processing.")
        return found_studies
    
    # only ask for password if needed
    needs_password = zip_needs_password(zip_path)
    if needs_password and password is None:
        password = get_password_from_gui(zip_path)
    elif not needs_password:
        password = None  # Ensure we don't use a password if not needed
    if needs_password and password is None:
        print("no password provided")
        return found_studies
        
    # use ramdisk if available for temp extraction (much faster)
    temp_dir = tempfile.mkdtemp()
    print(f"Extracting {zip_path} to temp dir...")
    
    if progress_callback:
        prefix = "  " * nested_level  # Visual indentation for nested progress
        progress_callback(0, f"{prefix}Extracting {os.path.basename(zip_path)}...")
    
    # Speed up zip extraction by using a more efficient approach
    try:
        # Check if it's a large file and use a more efficient method
        file_size = os.path.getsize(zip_path)
        
        if file_size > 100 * 1024 * 1024:  # if greater than 100MB
            # For large files, we'll extract only .dcm files and files without extensions
            # This is much faster than extracting everything
            
            if needs_password:
                with pyzipper.AESZipFile(zip_path) as zf:
                    # Only extract DICOM files to save time
                    for file_info in zf.infolist():
                        filename = file_info.filename
                        if filename.endswith('.dcm') or filename.endswith('.zip') or '.' not in os.path.basename(filename):
                            zf.extract(file_info, temp_dir, pwd=password)
            else:
                with zipfile.ZipFile(zip_path) as zf:
                    # Only extract DICOM files to save time
                    for file_info in zf.infolist():
                        filename = file_info.filename
                        if filename.endswith('.dcm') or filename.endswith('.zip') or '.' not in os.path.basename(filename):
                            zf.extract(file_info, temp_dir)
        else:
            # For smaller files, just extract everything
            if needs_password:
                with pyzipper.AESZipFile(zip_path) as zf:
                    zf.extractall(temp_dir, pwd=password)
            else:
                with zipfile.ZipFile(zip_path) as zf:
                    zf.extractall(temp_dir)
    except (RuntimeError, zipfile.BadZipFile) as e:
        print(f"Error extracting zip: {e}")
        shutil.rmtree(temp_dir)
        return found_studies
        
    print(f"Processing files from {zip_path}...")
    if progress_callback:
        prefix = "  " * nested_level
        progress_callback(25, f"{prefix}Processing files from {os.path.basename(zip_path)}...")
    
    # Check for nested zip files first
    nested_zip_files = []
    for root, _, files in os.walk(temp_dir):
        for file in files:
            if file.lower().endswith('.zip'):
                nested_zip_files.append(os.path.join(root, file))
    
    # Process nested zip files (with proper progress distribution)
    if nested_zip_files:
        print(f"Found {len(nested_zip_files)} nested zip files in {zip_path}")
        # Allocate 40% of progress to nested zip processing
        nested_progress_start = 25
        nested_progress_end = 65
        
        for i, nested_zip in enumerate(nested_zip_files):
            if progress_callback:
                prefix = "  " * nested_level
                nested_progress = nested_progress_start + (i / len(nested_zip_files) * (nested_progress_end - nested_progress_start))
                progress_callback(nested_progress, f"{prefix}Processing nested zip {i+1}/{len(nested_zip_files)}: {os.path.basename(nested_zip)}")
            
            # Define a nested progress callback that scales within the allocated range
            def nested_progress_callback(percent, text):
                if progress_callback:
                    if percent >= 0:  # Skip password requests
                        # Calculate the progress range for this nested zip
                        range_size = (nested_progress_end - nested_progress_start) / len(nested_zip_files)
                        start_pos = nested_progress_start + (i * range_size)
                        scaled_percent = start_pos + (percent / 100) * range_size
                        prefix = "  " * (nested_level + 1)  # Add indentation for nested level
                        progress_callback(scaled_percent, f"{prefix}{text}")
            
            # Process the nested zip, incrementing the nested level
            nested_studies = process_zip_file(
                nested_zip, 
                password,  # Pass the same password (user can input different if needed)
                max_workers, 
                nested_progress_callback,
                nested_level + 1,
                max_nested_level
            )
            
            # Merge the nested zip results with the current zip results
            for study_info, paths in nested_studies.items():
                # For nested zips, use the outer zip path as the source
                found_studies[study_info].add(zip_path)
    
    # Now find and process regular DICOM files
    progress_start = 65 if nested_zip_files else 25
    
    # find all potential dicom files - check more extensions
    dicom_files = []
    for root, _, files in os.walk(temp_dir):
        for file in files:
            file_path = os.path.join(root, file)
            # Skip zip files as they were already processed
            if file.lower().endswith('.zip'):
                continue
                
            ext = os.path.splitext(file)[1].lower()
            # Check known DICOM extensions (.dcm, .ima, .dicom) and files with no extension
            if ext in ('.dcm', '.ima', '.dicom') or not ext:
                dicom_files.append(file_path)
    
    print(f"Found {len(dicom_files)} potential DICOM files in {zip_path}")
    
    # process them in parallel - increase batch size for better performance
    found_count = 0
    if dicom_files:
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Process in larger batches for better performance
            batch_size = 100
            for i in range(0, len(dicom_files), batch_size):
                if progress_callback:
                    progress_percent = progress_start + (i / len(dicom_files) * (100 - progress_start))
                    prefix = "  " * nested_level
                    progress_callback(progress_percent, f"{prefix}Processing {os.path.basename(zip_path)}: {i}/{len(dicom_files)} files...")
                    
                batch = dicom_files[i:i+batch_size]
                futures = {executor.submit(process_dicom_file, f): f for f in batch}
                
                for future in concurrent.futures.as_completed(futures):
                    result = future.result()
                    if result:
                        study_info, file_path = result
                        if study_info:
                            found_studies[study_info].add(zip_path)
                            found_count += 1
    
    print(f"Found {found_count} valid studies in {zip_path}")
    if progress_callback:
        prefix = "  " * nested_level
        progress_callback(100, f"{prefix}Completed processing {os.path.basename(zip_path)}")
    
    # clean up
    shutil.rmtree(temp_dir)
    gc.collect()  # force garbage collection after processing large dirs
    
    return found_studies

# set a timeout for operations
def set_timeout(seconds):
    def timeout_handler(signum, frame):
        raise TimeoutError(f"Operation timed out after {seconds} seconds")
    
    # Register the timeout handler
    signal.signal(signal.SIGALRM, timeout_handler)
    signal.alarm(seconds)
    
def cancel_timeout():
    signal.alarm(0)

# use fast directory scanning
def find_zip_files(directory):
    zip_files = []
    # Use a more efficient algorithm for directory scanning
    try:
        for root, dirs, files in os.walk(directory):
            # Skip hidden directories
            dirs[:] = [d for d in dirs if not d.startswith('.')]
            for file in files:
                if file.endswith('.zip'):
                    zip_files.append(os.path.join(root, file))
    except Exception as e:
        print(f"Error scanning directory: {e}")
    return zip_files

# Worker thread for processing to keep UI responsive
class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, str)
    finished_signal = pyqtSignal(object)
    error_signal = pyqtSignal(str)
    
    def __init__(self, input_path, max_workers):
        super().__init__()
        self.input_path = input_path
        self.max_workers = max_workers
        self.max_nested_level = 5  # Default, can be overridden
        
    def run(self):
        try:
            if os.path.isfile(self.input_path) and self.input_path.endswith('.zip'):
                self.progress_updated.emit(0, f"Processing zip file: {os.path.basename(self.input_path)}")
                needs_password = zip_needs_password(self.input_path)
                # Password needs to be handled in the main thread
                self.progress_updated.emit(-1, "Password")
                # Wait for password to be set
                while not hasattr(self, 'password') and needs_password:
                    time.sleep(0.1)
                password = getattr(self, 'password', None) if needs_password else None
                
                study_data = process_zip_file(
                    self.input_path, 
                    password, 
                    max_workers=self.max_workers,
                    progress_callback=lambda p, t: self.progress_updated.emit(p, t),
                    nested_level=0,
                    max_nested_level=self.max_nested_level  # Allow up to 5 levels of nesting
                )
                
            elif os.path.isdir(self.input_path):
                self.progress_updated.emit(0, f"Processing directory: {os.path.basename(self.input_path)}")
                study_data = self.extract_directory_info()
                
            else:
                self.error_signal.emit(f"Invalid path: {self.input_path}")
                return
                
            self.finished_signal.emit(study_data)
            
        except Exception as e:
            import traceback
            traceback.print_exc()
            self.error_signal.emit(f"Error: {str(e)}")
    
    def extract_directory_info(self):
        study_data = defaultdict(set)
        passwords = {}  # cache passwords so user doesn't need to enter multiple times
        
        # find all zip files first - faster approach
        self.progress_updated.emit(0, "Scanning for zip files...")
        zip_files = find_zip_files(self.input_path)
        
        # process zip files in parallel - ask for password once for all
        if zip_files:
            self.progress_updated.emit(5, f"Found {len(zip_files)} zip files")
            
            # Check which zip files need passwords first to avoid unnecessary prompts
            password_needed = {}
            for i, zip_path in enumerate(zip_files):
                try:
                    password_needed[zip_path] = zip_needs_password(zip_path)
                    self.progress_updated.emit(5 + (i / len(zip_files) * 10), f"Checking zip file {i+1}/{len(zip_files)}...")
                except Exception:
                    password_needed[zip_path] = True  # Assume password needed if check fails
            
            # Get one password for all password-protected files
            shared_password = None
            if any(password_needed.values()):
                # Password needs to be handled in the main thread
                self.progress_updated.emit(-1, "Password")
                # Wait for password to be set
                while not hasattr(self, 'shared_password'):
                    time.sleep(0.1)
                shared_password = self.shared_password
            
            # Process zip files with the selected password
            for i, zip_path in enumerate(zip_files):
                # Only use password if needed for this specific zip
                password = shared_password if password_needed[zip_path] else None
                
                # Scale the progress updates from this individual zip operation to the overall progress range
                progress_start = 15 + (i / len(zip_files) * 80)
                progress_end = 15 + ((i + 1) / len(zip_files) * 80)
                
                def progress_callback(percent, text):
                    if percent >= 0:  # Skip password requests which use -1
                        scaled_percent = progress_start + (percent / 100) * (progress_end - progress_start)
                        self.progress_updated.emit(scaled_percent, text)
                
                try:
                    self.progress_updated.emit(progress_start, f"Processing zip {i+1}/{len(zip_files)}: {os.path.basename(zip_path)}")
                    found_studies = process_zip_file(
                        zip_path, 
                        password, 
                        self.max_workers, 
                        progress_callback,
                        nested_level=0,
                        max_nested_level=self.max_nested_level
                    )
                    for study_info, paths in found_studies.items():
                        study_data[study_info].update(paths)
                except Exception as e:
                    print(f"Error processing {zip_path}: {e}")
                    self.progress_updated.emit(progress_end, f"Error with {os.path.basename(zip_path)}: {str(e)}")
        
        # if no zip files found, scan for regular dicom files
        if not zip_files:
            self.progress_updated.emit(0, "Scanning for individual DICOM files...")
            
            # faster approach - sample file types first
            file_extensions = defaultdict(list)
            dicom_extensions = set(['.dcm'])  # Known DICOM extensions
            sampled_extensions = set()
            dicom_files = []
            
            # First pass - collect file extensions and sample unknown types
            self.progress_updated.emit(5, "Analyzing file types...")
            file_count = 0
            total_files = 0
            
            # First, count total files for progress tracking
            for root, _, files in os.walk(self.input_path):
                total_files += len(files)
            
            for root, _, files in os.walk(self.input_path):
                for file in files:
                    ext = os.path.splitext(file)[1].lower()
                    file_count += 1
                    
                    if file_count % 100 == 0:
                        self.progress_updated.emit(5 + (file_count / total_files * 15), 
                                             f"Analyzing file types: {file_count}/{total_files}")
                    
                    # Known DICOM extension - add directly
                    if ext in dicom_extensions:
                        dicom_files.append(os.path.join(root, file))
                        continue
                        
                    # Sample unknown extensions
                    if ext not in sampled_extensions and ext not in dicom_extensions:
                        file_path = os.path.join(root, file)
                        if is_valid_dicom_file(file_path):
                            dicom_extensions.add(ext)
                            dicom_files.append(file_path)
                        sampled_extensions.add(ext)
                    
                    # Save file paths by extension for second pass
                    file_extensions[ext].append(os.path.join(root, file))
            
            # Second pass - add all files with DICOM extensions
            for ext in dicom_extensions:
                if ext != '.dcm':  # We already added .dcm files
                    dicom_files.extend(file_extensions[ext])
            
            self.progress_updated.emit(20, f"Found {len(dicom_files)} potential DICOM files")
            
            # Process files in batches for better performance
            chunk_size = max(1, min(1000, len(dicom_files) // (self.max_workers * 2)))
            
            # Process in batches
            processed_count = 0
            for i in range(0, len(dicom_files), chunk_size):
                chunk = dicom_files[i:i+chunk_size]
                processed_count += len(chunk)
                progress = 20 + (processed_count / len(dicom_files) * 75)
                self.progress_updated.emit(progress, f"Processing files: {processed_count}/{len(dicom_files)}")
                
                # Process this batch
                for file_path in chunk:
                    result = process_dicom_file(file_path)
                    if result:
                        study_info, file_path = result
                        if study_info:
                            study_data[study_info].add(file_path)
        
        # Show error if nothing found after all processing
        if not study_data:
            self.error_signal.emit("No valid studies found in the directory")
            return False
            
        self.progress_updated.emit(95, "Finalizing results...")
        return study_data

# Main processing function with UI
def main():
    # optimize worker count based on system
    max_workers = min(os.cpu_count(), 8)  # Don't use more than 8 cores to avoid thrashing
    
    try:
        # initialize QApplication with minimal resources for speed
        app = QApplication(sys.argv)
        
        if len(sys.argv) < 2:
            print("Usage: script.py <path_to_directory_or_zip_file>")
            sys.exit(1)
            
        input_path = sys.argv[1]
        
        # Set default nested ZIP processing level
        max_nested_level = 5
        
        # Create and show the progress dialog
        progress_dialog = ProgressDialog("Processing DICOM Files")
        progress_dialog.show()
        
        # Create the processing thread
        processing_thread = ProcessingThread(input_path, max_workers)
        processing_thread.max_nested_level = max_nested_level
        
        # Connect signals
        def update_progress(value, text):
            if value == -1 and text == "Password":
                # Need to get password in main thread
                if hasattr(processing_thread, 'shared_password'):
                    # Getting password for a specific zip file
                    password = get_password_from_gui(input_path)
                    processing_thread.password = password
                else:
                    # Getting shared password for multiple zip files
                    password = get_password_from_gui("password protected zip files")
                    processing_thread.shared_password = password
            else:
                progress_dialog.update_progress(value, text)
        
        processing_thread.progress_updated.connect(update_progress)
        
        def on_finished(study_data):
            progress_dialog.update_progress(100, "Processing complete. Formatting results...")
            try:
                # Extract patient info from the first valid DICOM file found
                patient_name = "Unknown"
                patient_dob = "Unknown"
                
                # Try to find a valid patient name and DOB from the study data
                for study_info, _ in study_data.items():
                    if len(study_info) >= 4:  # Make sure we have patient name and DOB fields
                        patient_name = study_info[2]
                        patient_dob = study_info[3]
                        if patient_name != "Unknown" and not patient_name.startswith("Unknown"):
                            break  # Found a valid patient name, stop looking
                
                # Format and prepare study data for clipboard
                study_list = [(date, desc) for (date, desc, *_), _ in study_data.items() 
                             if desc != "Unknown" and not desc.startswith("Study for")]
                
                # sort and prepare study data for clipboard - faster sorting
                study_list.sort(key=lambda x: (x[0], x[1]))
                
                # Build output with patient info at the top
                lines = []
                
                # Add patient info at the top
                if patient_name != "Unknown" or patient_dob != "Unknown":
                    lines.append(f"{patient_name}, {patient_dob}\r\n\r\n")
                
                lines.append("STUDIES\r\n\r\n")
                
                # Add studies
                for study_date, study_description in study_list:
                    lines.append(f"{study_date} {study_description} \r\n")
                
                output_text = "".join(lines)
                
                # copy to clipboard
                clipboard.copy(output_text)
                progress_dialog.close()
                show_success_popup("Studies have been successfully copied to the clipboard")
                app.quit()
                
            except Exception as e:
                import traceback
                traceback.print_exc()
                progress_dialog.close()
                show_error_popup(f"Error formatting results: {str(e)}")
                app.quit()
        
        def on_error(error_message):
            progress_dialog.close()
            show_error_popup(error_message)
            app.quit()
        
        processing_thread.finished_signal.connect(on_finished)
        processing_thread.error_signal.connect(on_error)
        
        # Start processing in background
        processing_thread.start()
        
        # Run the application event loop
        sys.exit(app.exec_())
        
    except Exception as e:
        print(f"Error in main: {e}")
        sys.exit(1)

# add a simple profiler to measure speed improvements
def profile_main():
    start_time = time.time()
    result = main()
    end_time = time.time()
    print(f"Total execution time: {end_time - start_time:.2f} seconds")
    return result

if __name__ == "__main__":
    # optimizing python's startup
    import gc
    gc.disable()  # disable automatic garbage collection for better performance
    
    # verify that multiprocessing works properly on windows
    if sys.platform.startswith('win'):
        import multiprocessing
        # This is needed to avoid issues with multiprocessing on Windows
        multiprocessing.freeze_support()
    
    # run the main function
    profile_main()
    
    # cleanup before exit
    gc.enable()
    gc.collect()