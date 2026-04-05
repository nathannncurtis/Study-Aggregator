import sys
import os
import subprocess
import base64
import json
import time
import logging
from pathlib import Path

import clipboard
from PyQt5.QtWidgets import (QApplication, QMessageBox, QInputDialog, QLineEdit, QProgressBar,
                              QLabel, QVBoxLayout, QWidget)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon


def _get_app_root():
    """Get the application root directory, handling COIL bundled layouts."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    app_dir = os.path.dirname(os.path.abspath(__file__))
    parent = os.path.dirname(os.path.dirname(app_dir))
    if os.path.basename(os.path.dirname(app_dir)) == '_internal':
        return parent
    return app_dir


APP_ROOT = _get_app_root()


# --- Logging ---

def setup_logging():
    try:
        log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(filename)s:%(lineno)d - %(message)s')
        logger = logging.getLogger()
        logger.setLevel(logging.DEBUG)

        for handler in logger.handlers[:]:
            if isinstance(handler, logging.StreamHandler) and handler.stream in [sys.stdout, sys.stderr]:
                logger.removeHandler(handler)

        try:
            log_file_path = os.path.join(APP_ROOT, "dicom_aggregator.log")
            has_file_handler = any(
                isinstance(h, logging.FileHandler) and h.baseFilename == os.path.abspath(log_file_path)
                for h in logger.handlers
            )
            if not has_file_handler:
                file_handler = logging.FileHandler(log_file_path, mode='a')
                file_handler.setFormatter(log_formatter)
                logger.addHandler(file_handler)
        except Exception:
            console_handler = logging.StreamHandler(sys.stderr)
            console_handler.setFormatter(log_formatter)
            logger.addHandler(console_handler)

        logging.info("=== DICOM Aggregator Session Started ===")
        logging.info(f"Running as frozen executable: {'Yes' if getattr(sys, 'frozen', False) else 'No'}")
        logging.info(f"Command line args: {sys.argv}")
        return True
    except Exception as e:
        print(f"Critical: Failed to setup logging: {e}", file=sys.stderr)
        return False


def handle_critical_error(error, context="Unknown"):
    error_msg = str(error)
    logging.critical(f"Critical error in {context}: {error_msg}", exc_info=True)

    if "No module named" in error_msg:
        return "Missing required software component. Please reinstall the application."
    elif "Permission denied" in error_msg or "Access is denied" in error_msg:
        return "Permission denied accessing files or directories. Check file permissions."
    elif "Memory" in error_msg:
        return "Insufficient memory to process files. Try processing smaller batches."
    else:
        return f"An unexpected error occurred: {error_msg}. Check the log file for details."


setup_logging()

# --- Icon ---

icon_path = os.path.join(APP_ROOT, 'agg.ico')


# --- Engine binary discovery ---

def find_engine_binary():
    """Locate the study-agg-engine binary."""
    candidates = [
        os.path.join(APP_ROOT, 'study-agg-engine.exe'),
        os.path.join(APP_ROOT, 'engine', 'target', 'release', 'study-agg-engine.exe'),
        os.path.join(APP_ROOT, 'engine', 'target', 'debug', 'study-agg-engine.exe'),
    ]
    for path in candidates:
        if os.path.isfile(path):
            logging.info(f"Found engine binary: {path}")
            return path

    logging.error("Engine binary not found. Searched: " + ", ".join(candidates))
    return None


# --- GUI Components ---

class ProgressDialog(QWidget):
    def __init__(self, title="Processing"):
        super().__init__()
        self.setWindowTitle(title)
        self.setWindowFlag(Qt.WindowStaysOnTopHint)
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        layout = QVBoxLayout()

        self.label = QLabel("Processing files...")
        layout.addWidget(self.label)

        self.progress_bar = QProgressBar()
        self.progress_bar.setMinimum(0)
        self.progress_bar.setMaximum(100)
        layout.addWidget(self.progress_bar)

        self.setLayout(layout)
        self.resize(400, 120)

    def update_progress(self, value, text=None):
        try:
            self.progress_bar.setValue(max(0, value))
            if text:
                self.label.setText(text)
        except Exception as e:
            logging.error(f"Error updating progress: {e}")

    def closeEvent(self, event):
        event.ignore()


def show_error_popup(message):
    logging.error(f"Displaying error popup: {message}")
    try:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Critical)
        msg.setText(message)
        msg.setWindowTitle("Study Aggregator - Error")
        if os.path.exists(icon_path):
            msg.setWindowIcon(QIcon(icon_path))
        msg.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg.exec_()
    except Exception as e:
        logging.error(f"Failed to show error popup: {e}")
        print(f"ERROR: {message}", file=sys.stderr)


def show_success_popup(message):
    logging.info(f"Displaying success popup: {message}")
    try:
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText(message)
        msg.setWindowTitle("Study Aggregator - Success")
        if os.path.exists(icon_path):
            msg.setWindowIcon(QIcon(icon_path))
        msg.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg.exec_()
    except Exception as e:
        logging.error(f"Failed to show success popup: {e}")


# --- Processing Thread (Rust engine subprocess) ---

class ProcessingThread(QThread):
    progress_updated = pyqtSignal(int, str)
    finished_signal = pyqtSignal(object)
    error_signal = pyqtSignal(str)
    set_busy = pyqtSignal()
    set_normal = pyqtSignal()

    def __init__(self, input_path, engine_path, progress_dialog=None):
        super().__init__()
        self.input_path = input_path
        self.engine_path = engine_path
        self.progress_dialog = progress_dialog
        self._password_b64 = None

    def run(self):
        try:
            self.set_busy.emit()
            self._run_engine(password_b64=None)
        except Exception as e:
            error_msg = handle_critical_error(e, "processing thread")
            self.error_signal.emit(error_msg)
        finally:
            self.set_normal.emit()

    def _run_engine(self, password_b64=None):
        cmd = [self.engine_path, self.input_path]

        if password_b64:
            cmd.extend(['--password', password_b64])

        logging.info(f"Running engine: {' '.join(cmd[:2])} ...")

        try:
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0),
            )
        except Exception as e:
            self.error_signal.emit(f"Failed to start engine: {e}")
            return

        engine_errors = []

        # Read stdout in a background thread to prevent pipe deadlock.
        # The engine writes progress JSON to stderr and the final result
        # JSON to stdout; if stdout fills its OS pipe buffer (64KB) before
        # we read it, the engine blocks and never closes stderr.
        import threading
        stdout_chunks = []
        def _drain_stdout():
            try:
                stdout_chunks.append(process.stdout.read())
            except Exception:
                pass
        stdout_thread = threading.Thread(target=_drain_stdout, daemon=True)
        stdout_thread.start()

        for raw_line in process.stderr:
            line = raw_line.decode('utf-8', errors='replace').strip()
            if not line:
                continue
            try:
                msg = json.loads(line)
                msg_type = msg.get('type')

                if msg_type == 'progress':
                    pct = msg.get('percent', 0)
                    text = msg.get('message', '')
                    self.progress_updated.emit(max(0, pct), text)

                elif msg_type == 'password_needed':
                    self.progress_updated.emit(-1, "Password")
                    wait_start = time.time()
                    while not hasattr(self, '_password_received') and (time.time() - wait_start < 600):
                        time.sleep(0.1)

                    if hasattr(self, '_password_received'):
                        pw_bytes = self._password_received
                        del self._password_received

                        if pw_bytes is None:
                            self.error_signal.emit("Password input cancelled.")
                            process.kill()
                            stdout_thread.join(timeout=2)
                            return

                        process.kill()
                        process.wait()
                        stdout_thread.join(timeout=2)
                        pw_b64 = base64.b64encode(pw_bytes).decode('ascii')
                        self._run_engine(password_b64=pw_b64)
                        return
                    else:
                        self.error_signal.emit("Password input timed out.")
                        process.kill()
                        stdout_thread.join(timeout=2)
                        return

                elif msg_type == 'error':
                    err_msg = msg.get('message', 'Unknown engine error')
                    logging.error(f"Engine error: {err_msg}")
                    engine_errors.append(err_msg)

            except json.JSONDecodeError:
                logging.debug(f"Engine stderr: {line}")

        stdout_thread.join(timeout=30)
        stdout_data = b''.join(stdout_chunks)
        exit_code = process.wait()

        if exit_code == 2 and not password_b64:
            self.progress_updated.emit(-1, "Password")
            wait_start = time.time()
            while not hasattr(self, '_password_received') and (time.time() - wait_start < 600):
                time.sleep(0.1)

            if hasattr(self, '_password_received'):
                pw_bytes = self._password_received
                del self._password_received
                if pw_bytes:
                    pw_b64 = base64.b64encode(pw_bytes).decode('ascii')
                    self._run_engine(password_b64=pw_b64)
                    return

            self.error_signal.emit("Password required but not provided.")
            return

        if exit_code != 0:
            if engine_errors:
                self.error_signal.emit("\n".join(engine_errors))
            else:
                self.error_signal.emit(
                    f"Processing failed (exit code {exit_code}). Check log for details."
                )
            return

        try:
            result = json.loads(stdout_data.decode('utf-8'))
        except (json.JSONDecodeError, UnicodeDecodeError) as e:
            self.error_signal.emit(f"Failed to parse engine output: {e}")
            return

        patients = result.get('patients', {})
        stats = result.get('stats', {})

        logging.info(
            f"Engine completed: {stats.get('files_scanned', 0)} files scanned, "
            f"{stats.get('dicom_valid', 0)} valid DICOM, "
            f"{stats.get('patients_found', 0)} patients, "
            f"{stats.get('elapsed_ms', 0)}ms"
        )

        if not patients:
            self.error_signal.emit(
                "No DICOM studies found in the specified location. "
                "Please verify that the source contains valid DICOM files."
            )
            return

        for patient in patients.values():
            for study in patient.get('studies', {}).values():
                study['all_series'] = set(study.get('all_series', []))

        self.finished_signal.emit(patients)


# --- Path Handling ---

def normalize_and_validate_path(input_path):
    try:
        normalized_path = os.path.normpath(input_path)

        if len(normalized_path) == 3 and normalized_path.endswith(':\\'):
            try:
                os.listdir(normalized_path)
                return normalized_path
            except (OSError, PermissionError):
                raise Exception(f"Cannot access drive {normalized_path}.")

        if not os.path.exists(normalized_path):
            raise Exception(f"The specified path does not exist: '{normalized_path}'")

        if os.path.isdir(normalized_path):
            try:
                os.listdir(normalized_path)
            except (OSError, PermissionError):
                raise Exception(f"Cannot access directory '{normalized_path}'. Check permissions.")

        return normalized_path
    except Exception as e:
        logging.error(f"Path validation failed: {e}")
        raise


def clean_input_path(raw_path):
    try:
        cleaned_path = raw_path.strip('"\'')
        if len(cleaned_path) == 2 and cleaned_path[1] == ':':
            cleaned_path = cleaned_path + '\\'
        return cleaned_path
    except Exception:
        return raw_path


def get_password_from_gui(parent_widget, description_text):
    try:
        password, ok = QInputDialog.getText(
            parent_widget, 'Password Required',
            f'Enter password for {description_text}:',
            QLineEdit.Password
        )
        if ok and password:
            return password.encode('utf-8')
        elif ok and not password:
            return b''
        return None
    except Exception as e:
        logging.error(f"Error getting password from GUI: {e}")
        return None


# --- Main Application ---

def main_app_logic():
    try:
        if getattr(sys, 'frozen', False):
            try:
                os.chdir(APP_ROOT)
            except Exception:
                pass

        engine_path = find_engine_binary()
        if not engine_path:
            show_error_popup(
                "Engine binary (study-agg-engine.exe) not found. "
                "Please reinstall the application."
            )
            return 1

        app = QApplication.instance() or QApplication(sys.argv)

        app.setStyle("Fusion")
        app.setStyleSheet("""
        QWidget {
            background-color: #202020;
            color: #ffffff;
            font-family: "Segoe UI", Arial, sans-serif;
            font-size: 9pt;
        }
        QLabel { color: #ffffff; background-color: transparent; }
        QProgressBar {
            border: 1px solid #3f3f3f; border-radius: 3px;
            text-align: center; background-color: #1a1a1a; color: #ffffff;
        }
        QProgressBar::chunk { background-color: #0078d4; border-radius: 2px; }
        QDialog { background-color: #202020; }
        QMessageBox { background-color: #202020; }
        QLineEdit {
            background-color: #2d2d2d; color: #ffffff;
            border: 1px solid #3f3f3f; padding: 4px; border-radius: 3px;
        }
        QLineEdit:focus { border: 1px solid #0078d4; }
        """)

        raw_input_path = None
        for arg in sys.argv:
            cleaned = clean_input_path(arg)
            if os.path.exists(cleaned):
                raw_input_path = cleaned
                break

        if raw_input_path is None:
            show_error_popup(
                "No input specified. Please drag and drop a directory or ZIP file "
                "onto this application, or run it from the command line with a path argument."
            )
            return 1

        input_path = raw_input_path
        logging.info(f"Input path: '{input_path}'")

        try:
            validated_path = normalize_and_validate_path(input_path)
        except Exception as path_error:
            show_error_popup(str(path_error))
            return 1

        progress_dialog = ProgressDialog("Processing DICOM Files")
        progress_dialog.show()

        processing_thread = ProcessingThread(validated_path, engine_path, progress_dialog)

        def update_progress_slot(value, text):
            try:
                if value == -1 and text == "Password":
                    progress_dialog.hide()
                    try:
                        if os.path.isfile(validated_path) and validated_path.lower().endswith('.zip'):
                            pw = get_password_from_gui(
                                progress_dialog,
                                f"encrypted ZIP: {os.path.basename(validated_path)}"
                            )
                        else:
                            pw = get_password_from_gui(
                                progress_dialog,
                                "any password-protected ZIP files in the directory"
                            )
                        processing_thread._password_received = pw
                    finally:
                        progress_dialog.show()
                else:
                    progress_dialog.update_progress(value, text)
            except Exception as e:
                logging.error(f"Error updating progress: {e}")

        processing_thread.progress_updated.connect(update_progress_slot)
        processing_thread.set_busy.connect(lambda: app.setOverrideCursor(Qt.WaitCursor))
        processing_thread.set_normal.connect(lambda: app.restoreOverrideCursor())

        def on_finished_slot(merged_patients):
            logging.debug(f"Received {len(merged_patients)} patients from engine")
            try:
                progress_dialog.update_progress(100, "Processing complete. Formatting results...")

                if not merged_patients:
                    show_error_popup(
                        "No patient data found after processing. "
                        "Please verify the source contains valid DICOM files."
                    )
                    progress_dialog.close()
                    app.quit()
                    return

                def sort_patients(patient_data):
                    pid = patient_data.get('patient_id') or 'ZZZZ'
                    name = patient_data.get('patient_name', 'Unknown')
                    pid_val = pid.zfill(10) if pid.isdigit() else str(pid) if pid else 'ZZZZ'
                    return (pid_val, name.lower())

                sorted_patients = sorted(merged_patients.values(), key=sort_patients)
                lines = []

                for patient in sorted_patients:
                    pid = patient.get('patient_id', '')
                    name = patient.get('patient_name', 'Unknown')
                    dob = patient.get('patient_dob', 'Unknown')

                    display_name = f"NAME: {name} DOB: {dob}, ID: {pid or 'Unknown'}"
                    lines.extend([f"{display_name}\r\n", "STUDIES\r\n\r\n"])

                    studies_dict = patient.get('studies', {})
                    if isinstance(studies_dict, dict) and studies_dict:
                        sorted_studies = sorted(
                            studies_dict.values(),
                            key=lambda x: (x.get('study_date', 'Unknown'), x.get('study_description', ''))
                        )
                        for study in sorted_studies:
                            study_date = study.get('study_date', 'Unknown')
                            study_desc = study.get('study_description', 'Unknown')
                            all_series = study.get('all_series', set())
                            series_count = len(all_series)
                            if series_count > 0:
                                lines.append(f"{study_date} {study_desc} ({series_count} series)\r\n")
                            else:
                                lines.append(f"{study_date} {study_desc}\r\n")

                    lines.append("\r\n" + "=" * 50 + "\r\n\r\n")

                output_text = "".join(lines) if lines else None

                progress_dialog.close()

                if output_text:
                    try:
                        clipboard.copy(output_text)
                        show_success_popup(
                            "DICOM study information has been copied to your clipboard."
                        )
                    except Exception as e:
                        logging.error(f"Clipboard error: {e}")
                        show_error_popup(f"Failed to copy to clipboard: {e}")
                else:
                    show_error_popup("No study information to copy.")

                app.quit()

            except Exception as e:
                error_msg = handle_critical_error(e, "result formatting")
                progress_dialog.close()
                show_error_popup(f"Error formatting results: {error_msg}")
                app.quit()

        def on_error_slot(error_message):
            logging.error(f"Processing error: {error_message}")
            progress_dialog.close()
            show_error_popup(error_message)
            app.quit()

        processing_thread.finished_signal.connect(on_finished_slot)
        processing_thread.error_signal.connect(on_error_slot)

        logging.info(f"Starting engine for: {validated_path}")
        processing_thread.start()
        return app.exec_()

    except Exception as e:
        error_msg = handle_critical_error(e, "application startup")
        try:
            show_error_popup(f"Critical application error: {error_msg}")
        except Exception:
            print(f"CRITICAL ERROR: {error_msg}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    try:
        if sys.platform.startswith('win'):
            import multiprocessing
            multiprocessing.freeze_support()
        sys.exit(main_app_logic())
    except Exception as e:
        print(f"Fatal application error: {e}", file=sys.stderr)
        sys.exit(1)
