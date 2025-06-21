import sys
import os

# Attempt to set TCL/TK library paths if running frozen (e.g., PyInstaller)
# This might not be strictly necessary for a PySide6 app but can prevent
# issues if any lingering Tkinter-related imports or checks exist.
if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS  # PyInstaller's temporary folder
    tcl_library_path = os.path.join(base_dir, "tcl", "tcl8.6")
    tk_library_path = os.path.join(base_dir, "tcl", "tk8.6")
    if os.path.exists(tcl_library_path):
        os.environ["TCL_LIBRARY"] = tcl_library_path
    if os.path.exists(tk_library_path):
        os.environ["TK_LIBRARY"] = tk_library_path

from PySide6 import QtWidgets, QtCore, QtGui

# Placeholder for translations - will be properly integrated later
from translations import set_language, get_translation
import translations # For CURRENT_LANGUAGE access

# Placeholder for other imports that will be needed
import sqlite3
import re
import tempfile
import shutil
import platform
import subprocess
import logging
import datetime
# queue module won't be directly used for UI updates, but QObject workers will replace it.
# from concurrent.futures import ThreadPoolExecutor # Will be replaced by QThreadPool
import json # Added for user data persistence if needed, though db is primary
from passlib.hash import pbkdf2_sha256
from controllers.user_controller import UserController
from controllers.archive_controller import ArchiveController


# --- Global user store (loaded from DB) ---
# This will be populated by load_users_from_db
users = {}

# --- Constants (copied from original, may need review) ---
DEFAULT_ADMIN_PASSWORD = "admin123"
DEFAULT_USER_PASSWORD = "user123"
IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff"]
DOCUMENT_EXTENSIONS = [".xlsx", ".xls", ".doc", ".docx", ".ppt", ".pptx", ".pdf"]
SUPPORTED_FILE_EXTENSIONS = IMAGE_EXTENSIONS + DOCUMENT_EXTENSIONS
# set_language("en") # Initial language setting

# --- Logging Setup (copied from original, needs to be callable) ---
class LiveWritingFileHandler(logging.FileHandler):
    def emit(self, record):
        super().emit(record)
        self.flush()

def get_data_dir():
    if getattr(sys, 'frozen', False):
        return os.path.join(os.getenv('APPDATA'), 'FileArchiveApp')
    else:
        # For development, place it where pyside_main.py is
        return os.path.dirname(os.path.abspath(__file__))

def setup_logging():
    data_dir = get_data_dir()
    os.makedirs(data_dir, exist_ok=True)
    log_file_path = os.path.join(data_dir, 'archive_app_pyside.log') # New log file name

    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(module)s - %(funcName)s - %(message)s')
    handler = LiveWritingFileHandler(log_file_path)
    handler.setFormatter(log_formatter)
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)

    if logger.hasHandlers():
        logger.handlers.clear()
    logger.addHandler(handler)

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)
    logging.info(f"--- Logging initialized (PySide6). Log file at: {log_file_path} ---")

# Call logging setup at the start
setup_logging()


class FileArchiveApp(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        logging.info("FileArchiveApp (PySide6) __init__ started.")

        # Set initial language (can be moved or made configurable)
        set_language("en") # Default language

        self.current_user = None # Will be set after login
        self.current_company = {} # To store current company info {"display_name": "...", "safe_name": "..."}

        # --- Database and Controllers ---
        # The global `users` dict will be populated by initialize_user_database
        self.initialize_user_database() # Connects and loads users into global `users` dict

        # Ensure users dict is loaded before UserController initialization
        # self.load_users_from_db() # This is called within initialize_user_database

        self.user_controller = UserController(users) # Uses the global `users` dict

        # --- Paths and State ---
        self.archives_path = "archives"
        try:
             os.makedirs(self.archives_path, exist_ok=True)
        except OSError as e:
             logging.error(f"CRITICAL: Could not create archives directory '{self.archives_path}'. Error: {e}", exc_info=True)
             # In a Qt app, show a QMessageBox and then exit
             critical_error_box = QtWidgets.QMessageBox()
             critical_error_box.setIcon(QtWidgets.QMessageBox.Critical)
             critical_error_box.setWindowTitle("Fatal Error")
             critical_error_box.setText(f"Could not create required directory:\n{self.archives_path}\n\n{e}\n\nApplication cannot continue.")
             critical_error_box.exec()
             sys.exit(1)

        self.search_queries = [] # For dashboard analytics
        self.search_queries_lock = QtCore.QMutex() # For thread-safe access to search_queries

        # --- Document Structure Definition (copied from original) ---
        self.structure = {
            "Permanent Audit File": {
                "c1": {}, "c2": {}, "c3": {}, "c4": {}, "c5": {}, "c6": {},
            },
            "Working Papers File": {
                "A": {str(i): [] for i in range(1, 16)},
                "B1": {
                    "B10": ["B10A"], "B11": ["B11A"], "B12": [], "B13": [], "B14": [], "B15": [],
                    "B16": [], "B17": [], "B18": [], "B19": [],
                },
                "B2": {
                    "B20": [], "B21": [], "B22": [], "B23": [], "B24": [],
                    "B25": [], "B26": [], "B27": [],
                },
                "B3": {
                    "B30": [], "B31": [], "B32": [], "B33": [], "B34": [],
                },
                "I1": {
                     "I10": [], "I11": [], "I11A": [], "I12": [],
                 },
                 "I2": {
                     "I20": [], "I21": [], "I22": [], "I23": [], "I24": [],
                     "I25": [], "I26": [], "I27": [], "I28": [], "I29": [],
                 }
            }
        }
        # Create ArchiveController *after* self.structure is defined
        self.archive_controller = ArchiveController(self.structure, self.archives_path)

        # Initialize QThreadPool for background tasks
        self.thread_pool = QtCore.QThreadPool()
        # Consider setting maxThreadCount if needed, default is usually fine
        logging.info(f"QThreadPool initialized with maxThreadCount: {self.thread_pool.maxThreadCount()}")

        self.setWindowTitle(self._t("app_title"))

# --- Qt Worker Class for Threading ---
class Worker(QtCore.QObject):
    """
    Generic worker for running functions in a separate thread.
    """
    finished = QtCore.Signal()
    error = QtCore.Signal(str) # Emits error message string
    result = QtCore.Signal(object) # Emits any result object from the task
    progress = QtCore.Signal(int) # Emits progress percentage (0-100)

    def __init__(self, fn, *args, **kwargs):
        super().__init__()
        self.fn = fn
        self.args = args
        self.kwargs = kwargs
        self.is_running = True

    @QtCore.Slot()
    def run(self):
        try:
            if not self.is_running: # Should not happen if used correctly with QThreadPool
                return
            res = self.fn(*self.args, **self.kwargs)
            if self.is_running: # Check again before emitting, in case of early stop
                self.result.emit(res)
        except Exception as e:
            logging.error(f"Error in worker thread executing {self.fn.__name__}: {e}", exc_info=True)
            if self.is_running:
                self.error.emit(str(e))
        finally:
            if self.is_running:
                self.finished.emit()

    def stop(self): # Optional: way to signal worker to stop early if task supports it
        self.is_running = False
        self.setGeometry(100, 100, 950, 700) # x, y, width, height

        # Central widget and main layout
        self.central_widget = QtWidgets.QWidget()
        self.setCentralWidget(self.central_widget)
        self.main_layout = QtWidgets.QVBoxLayout(self.central_widget)
        self.main_layout.setContentsMargins(0,0,0,0) # Use full space

        # Placeholder for UI elements
        # self.setup_ui() # This will be called later

        logging.info("FileArchiveApp (PySide6) __init__ finished.")

    def _t(self, key):
        """Helper for translations."""
        return get_translation(key)

    def setup_ui(self):
        """Sets up the main UI components like header, tab widget, status bar."""
        logging.info("Setting up main UI elements.")

        # --- Header ---
        self.header_widget = QtWidgets.QWidget()
        self.header_layout = QtWidgets.QHBoxLayout(self.header_widget)
        self.header_layout.setContentsMargins(10, 5, 10, 5) # Add some padding

        self.app_title_label = QtWidgets.QLabel(self._t("ctklabel_text_file_archiving_system"))
        font = self.app_title_label.font()
        font.setPointSize(18) # Slightly larger
        font.setBold(True)
        self.app_title_label.setFont(font)
        self.header_layout.addWidget(self.app_title_label, stretch=1) # Stretch to push others right

        self.user_info_label = QtWidgets.QLabel("") # Placeholder, updated on login
        font = self.user_info_label.font()
        font.setPointSize(10)
        self.user_info_label.setFont(font)
        self.user_info_label.setAlignment(QtCore.Qt.AlignRight | QtCore.Qt.AlignVCenter)
        self.header_layout.addWidget(self.user_info_label, stretch=0)

        self.language_button = QtWidgets.QPushButton(self._t("button_text_switch_language"))
        self.language_button.setFixedWidth(120) # Fixed width for consistency
        self.language_button.clicked.connect(self.switch_language_qt)
        self.header_layout.addWidget(self.language_button, stretch=0)

        # Logout button will be added to a placeholder after login
        self.logout_button_placeholder_widget = QtWidgets.QWidget()
        self.logout_button_placeholder_layout = QtWidgets.QHBoxLayout(self.logout_button_placeholder_widget)
        self.logout_button_placeholder_layout.setContentsMargins(0,0,0,0)
        self.logout_button_placeholder_widget.setFixedWidth(110) # Space for logout button
        self.header_layout.addWidget(self.logout_button_placeholder_widget, stretch=0)

        self.main_layout.addWidget(self.header_widget)

        # --- TabWidget ---
        self.tab_widget = QtWidgets.QTabWidget()
        # self.tab_widget.setDocumentMode(True) # Optional: for a more modern look on some styles
        self.main_layout.addWidget(self.tab_widget, stretch=1) # Stretch to fill available space

        # --- Status Bar ---
        self.status_label = QtWidgets.QLabel(self._t("ctklabel_text_ready"))
        self.statusBar().addWidget(self.status_label, stretch=1)

        self.progress_bar = QtWidgets.QProgressBar()
        self.progress_bar.setRange(0,100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedWidth(150)
        self.progress_bar.setTextVisible(False)
        self.progress_bar.setVisible(False) # Initially hidden
        self.statusBar().addPermanentWidget(self.progress_bar)

        logging.info("Main UI elements (header, tabs, statusbar) setup.")

    def _update_header_texts(self):
        """Updates texts in the header, typically after language switch."""
        if hasattr(self, 'app_title_label'): # Check if UI elements exist
            self.app_title_label.setText(self._t("ctklabel_text_file_archiving_system"))
        if hasattr(self, 'language_button'):
            self.language_button.setText(self._t("button_text_switch_language"))

        if self.current_user and hasattr(self, 'user_info_label'):
            user_text = f"{self._t('label_user')}: {self.current_user['username']} ({self._t('role_' + self.current_user['role'])})"
            self.user_info_label.setText(user_text)
            self.setWindowTitle(f"{self._t('app_title')} - {self.current_user['username']}")
            if hasattr(self, 'logout_button') and self.logout_button: # Check if logout button exists
                self.logout_button.setText(self._t("ctkbutton_text_logout"))
        elif hasattr(self, 'user_info_label'):
            self.user_info_label.setText("")
            self.setWindowTitle(self._t("app_title"))


    def switch_language_qt(self):
        """Switches the application language and triggers UI rebuild."""
        current_lang = translations.CURRENT_LANGUAGE
        new_lang = "ar" if current_lang == "en" else "en"
        logging.info(f"Attempting to switch language from '{current_lang}' to '{new_lang}'")

        current_tab_idx = -1
        if hasattr(self, 'tab_widget'): # Ensure tab_widget exists
            current_tab_idx = self.tab_widget.currentIndex()

        if translations.set_language(new_lang):
            qapp = QtWidgets.QApplication.instance()
            if new_lang == "ar":
                qapp.setLayoutDirection(QtCore.Qt.RightToLeft)
            else:
                qapp.setLayoutDirection(QtCore.Qt.LeftToRight)

            self._rebuild_ui_for_language_qt(current_tab_idx)
            logging.info(f"Language switched to {new_lang} and UI rebuild initiated.")
        else:
            logging.error(f"Failed to set language to {new_lang}.")
            # Use QMessageBox for error display
            error_box = QtWidgets.QMessageBox(self)
            error_box.setIcon(QtWidgets.QMessageBox.Critical)
            error_box.setWindowTitle(self._t("error_title"))
            error_box.setText(self._t("error_language_switch_failed", lang=new_lang))
            error_box.exec()


    def _rebuild_ui_for_language_qt(self, selected_tab_idx=-1):
        """Rebuilds or updates UI elements for the new language."""
        logging.info(f"Rebuilding UI for language: {translations.CURRENT_LANGUAGE}")

        self._update_header_texts()

        if hasattr(self, 'status_label'):
            self.status_label.setText(self._t("ctklabel_text_ready"))
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("ctklabel_text_ready"), 2000)

        # Placeholder for full tab recreation logic
        # This will be expanded significantly in Step 5 and subsequent tab-specific steps.
        # For now, it primarily focuses on what's already built (header, status).
        if hasattr(self, 'tab_widget'):
            # Update titles of existing tabs (a more robust solution will replace tabs)
            # This is a simplified approach for now.
            # Example: if self.upload_tab: self.tab_widget.setTabText(self.tab_widget.indexOf(self.upload_tab), self._t("tab_upload_files"))

            # A more robust approach would be to call a method that re-initializes all tabs:
            if hasattr(self, 'setup_tabs_post_login'): # This method will be created in step 5
                 self.setup_tabs_post_login() # This should clear and re-add tabs with new translations

            if selected_tab_idx != -1 and selected_tab_idx < self.tab_widget.count():
                self.tab_widget.setCurrentIndex(selected_tab_idx)

        logging.info("UI text elements updated for new language.")

    # --- Tab Setup Methods ---
    def _create_placeholder_tab(self, title_key):
        """Helper to create a basic tab page with a title."""
        page = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout(page)
        label = QtWidgets.QLabel(f"Content for {self._t(title_key)}")
        label.setAlignment(QtCore.Qt.AlignCenter)
        font = label.font()
        font.setPointSize(14)
        label.setFont(font)
        layout.addWidget(label)
        return page

    def setup_upload_tab_qt(self):
        if not hasattr(self, 'tab_widget'): return
        self.upload_tab_page = self._create_placeholder_tab("tab_upload_files")
        self.tab_widget.addTab(self.upload_tab_page, self._t("tab_upload_files"))
        logging.info("Upload tab setup.")
        # Actual content for this tab will be in Step 8
        self.upload_tab_page.setObjectName("UploadTab") # For styling if needed

        main_upload_layout = QtWidgets.QVBoxLayout(self.upload_tab_page)

        scroll_area = QtWidgets.QScrollArea()
        scroll_area.setWidgetResizable(True)
        main_upload_layout.addWidget(scroll_area)

        scroll_content_widget = QtWidgets.QWidget()
        scroll_area.setWidget(scroll_content_widget)

        upload_content_layout = QtWidgets.QVBoxLayout(scroll_content_widget) # Layout for all content in scroll area

        # --- Company Info Group ---
        company_group_box = QtWidgets.QGroupBox(self._t("group_title_company_info"))
        company_layout = QtWidgets.QFormLayout(company_group_box)

        self.company_name_edit = QtWidgets.QLineEdit()
        self.company_name_edit.setPlaceholderText(self._t("placeholder_enter_company_name"))
        # Connect signal for when editing is finished or text changes significantly
        self.company_name_edit.editingFinished.connect(self._qt_trigger_structure_update_from_company)
        company_layout.addRow(QtWidgets.QLabel(self._t("label_company_name") + ":"), self.company_name_edit)
        upload_content_layout.addWidget(company_group_box)

        # --- Document Structure Group ---
        structure_group_box = QtWidgets.QGroupBox(self._t("group_title_document_structure"))
        structure_layout = QtWidgets.QFormLayout(structure_group_box)

        self.header_combo = QtWidgets.QComboBox()
        self.subheader_combo = QtWidgets.QComboBox()
        self.section_combo = QtWidgets.QComboBox()
        self.subsection_combo = QtWidgets.QComboBox()

        self.header_combo.addItems(list(self.structure.keys()))

        structure_layout.addRow(self._t("label_header") + ":", self.header_combo)
        structure_layout.addRow(self._t("label_subheader") + ":", self.subheader_combo)
        structure_layout.addRow(self._t("label_section") + ":", self.section_combo)
        structure_layout.addRow(self._t("label_subsection") + ":", self.subsection_combo)

        # Connect signals for combobox changes
        self.header_combo.currentIndexChanged.connect(self._qt_update_subheader_options)
        self.subheader_combo.currentIndexChanged.connect(self._qt_update_section_options)
        self.section_combo.currentIndexChanged.connect(self._qt_update_subsection_options)
        # subsection_combo usually doesn't trigger further updates downwards in this model

        # "Add New Folder Here..." button (Admin Only)
        if self.current_user and self.current_user["role"] == "admin":
            self.add_folder_button_upload_tab = QtWidgets.QPushButton(self._t("button_add_new_folder_here"))
            # self.add_folder_button_upload_tab.clicked.connect(self._open_add_folder_dialog_qt) # Dialog to be made
            structure_layout.addRow(self.add_folder_button_upload_tab)


        upload_content_layout.addWidget(structure_group_box)
        self._qt_update_subheader_options() # Initial population based on default header

        # --- Upload Options Group ---
        upload_options_group_box = QtWidgets.QGroupBox(self._t("group_title_upload_options"))
        upload_options_layout = QtWidgets.QHBoxLayout(upload_options_group_box) # Horizontal for buttons

        self.upload_file_button = QtWidgets.QPushButton(self._t("button_upload_file"))
        # self.upload_file_button.clicked.connect(self._upload_file_qt) # To be implemented
        upload_options_layout.addWidget(self.upload_file_button)

        self.batch_upload_button = QtWidgets.QPushButton(self._t("button_batch_upload"))
        # self.batch_upload_button.clicked.connect(self._batch_upload_qt) # To be implemented
        upload_options_layout.addWidget(self.batch_upload_button)

        self.scan_archive_button = QtWidgets.QPushButton(self._t("button_scan_archive"))
        # self.scan_archive_button.clicked.connect(self._scan_and_archive_qt) # To be implemented
        upload_options_layout.addWidget(self.scan_archive_button)
        upload_content_layout.addWidget(upload_options_group_box)

        # --- Drag & Drop Zone ---
        # This will be a custom widget (DropZoneWidget)
        self.drop_zone = DropZoneWidget(self) # Pass parent
        upload_content_layout.addWidget(self.drop_zone, stretch=1) # Stretch to take available space

        # Connect drop zone's signal to a handler in FileArchiveApp
        self.drop_zone.filesDropped.connect(self._process_dropped_files_qt)


    # --- Qt Specific Update Methods for ComboBoxes ---
    def _qt_trigger_structure_update_from_company(self):
        """Called when company name editing is finished."""
        company_name = self.company_name_edit.text().strip()
        if not company_name:
            # Optionally clear lower comboboxes or show a message
            self.header_combo.setCurrentIndex(0) # Reset header combo
            self._qt_update_subheader_options() # This will cascade updates
            return

        # If company structure needs to be created/verified on company name change:
        try:
            self.create_company_structure(company_name) # Ensures self.current_company is set
        except Exception as e:
            # Handle error (e.g., show QMessageBox)
            QtWidgets.QMessageBox.critical(self, self._t("error_title"),
                                           self._t("error_create_company_structure", company=company_name, error=str(e)))
            return

        # Refresh options based on the (potentially new) company structure
        self._qt_update_subheader_options()


    def _qt_update_subheader_options(self):
        company_name = self.company_name_edit.text().strip()
        header_key = self.header_combo.currentText()
        self.subheader_combo.clear()
        self.section_combo.clear()
        self.subsection_combo.clear()

        if not company_name or not header_key:
            return

        safe_company_name = self.sanitize_path(company_name)
        # Ensure current_company is updated for get_dynamic_folder_options if it relies on it
        if not hasattr(self, 'current_company') or self.current_company.get("safe_name") != safe_company_name:
            try:
                self.create_company_structure(company_name) # This sets self.current_company
                safe_company_name = self.current_company["safe_name"]
            except Exception: # If structure creation fails, no options to show
                return

        header_template_options = self.structure.get(header_key, {})
        company_header_path = os.path.join(self.archives_path, safe_company_name, header_key)

        subh_options = self.archive_controller.get_dynamic_folder_options(company_header_path, header_template_options)
        if subh_options:
            self.subheader_combo.addItems(subh_options)
        self._qt_update_section_options() # Cascade

    def _qt_update_section_options(self):
        company_name = self.company_name_edit.text().strip()
        header_key = self.header_combo.currentText()
        subheader_key = self.subheader_combo.currentText()
        self.section_combo.clear()
        self.subsection_combo.clear()

        if not company_name or not header_key or not subheader_key:
            return

        safe_company_name = self.sanitize_path(company_name)
        header_template = self.structure.get(header_key, {})
        section_template_options = None
        if isinstance(header_template, dict):
            section_template_options = header_template.get(subheader_key, None)

        company_subheader_path = os.path.join(self.archives_path, safe_company_name, header_key, subheader_key)
        sec_options = self.archive_controller.get_dynamic_folder_options(company_subheader_path, section_template_options)
        if sec_options:
            self.section_combo.addItems(sec_options)
        self._qt_update_subsection_options() # Cascade

    def _qt_update_subsection_options(self):
        company_name = self.company_name_edit.text().strip()
        header_key = self.header_combo.currentText()
        subheader_key = self.subheader_combo.currentText()
        section_key = self.section_combo.currentText()
        self.subsection_combo.clear()

        if not company_name or not header_key or not subheader_key or not section_key:
            return

        safe_company_name = self.sanitize_path(company_name)
        subsection_template_options = None
        header_template = self.structure.get(header_key, {})
        if isinstance(header_template, dict):
            section_dict = header_template.get(subheader_key, {})
            if isinstance(section_dict, dict):
                subsection_template_options = section_dict.get(section_key, None)

        company_section_path = os.path.join(self.archives_path, safe_company_name, header_key, subheader_key, section_key)
        sub_sec_options = self.archive_controller.get_dynamic_folder_options(company_section_path, subsection_template_options)
        if sub_sec_options:
            self.subsection_combo.addItems(sub_sec_options)

    def _process_dropped_files_qt(self, file_paths):
        """Handles files dropped onto the DropZoneWidget."""
        logging.info(f"Files dropped: {file_paths}")
        company_name = self.company_name_edit.text().strip()
        if not company_name:
            QtWidgets.QMessageBox.warning(self, self._t("error_title"), self._t("error_company_name_required_drop"))
            return

        # Gather structure selections
        header = self.header_combo.currentText()
        subheader = self.subheader_combo.currentText()
        section = self.section_combo.currentText()
        subsection = self.subsection_combo.currentText()

        # Filter for supported files (already done by DropZoneWidget, but good for safety)
        valid_files = [f for f in file_paths if os.path.splitext(f)[1].lower() in SUPPORTED_FILE_EXTENSIONS]
        if not valid_files:
            QtWidgets.QMessageBox.warning(self, self._t("error_title_invalid_files"), self._t("error_no_supported_files_dropped"))
            return

        # For drag and drop, we typically don't do interactive renaming per file.
        # The original Tkinter version had a strict check. Let's assume auto-rename or pre-check might be desired.
        # For now, let's use the batch_upload_qt logic which has auto-rename confirmation.
        # We'll set auto_rename_confirmed to True for simplicity in drag-drop, or prompt once.

        # Prompt once for the batch of dropped files if renaming might be needed.
        # This requires checking all files first against the current structure.
        # For now, let's assume auto_rename is implicitly true for drag-drop for simplicity of this step
        # A more complex UI might ask once for the whole dropped batch.

        self.batch_upload_qt(company_name, header, subheader, section, subsection, valid_files, auto_rename_confirmed=True, calling_tab_ref=self.upload_tab_page)


    def setup_manage_tab_qt(self):
        if not hasattr(self, 'tab_widget'): return
        self.manage_tab_page = self._create_placeholder_tab("tab_manage_files") # Keep placeholder for now
        self.tab_widget.addTab(self.manage_tab_page, self._t("tab_manage_files"))
        logging.info("Manage tab setup (placeholder).")
        # Actual content for this tab will be in Step 9

    def setup_settings_tab_qt(self):
        if not hasattr(self, 'tab_widget'): return
        self.settings_tab_page = self._create_placeholder_tab("tab_settings")
        self.tab_widget.addTab(self.settings_tab_page, self._t("tab_settings"))
        logging.info("Settings tab setup.")
        # Actual content for this tab will be in Step 10

    def setup_admin_tab_qt(self):
        if not hasattr(self, 'tab_widget'): return
        if self.current_user and self.current_user["role"] == "admin":
            self.admin_tab_page = self._create_placeholder_tab("tab_admin")
            self.tab_widget.addTab(self.admin_tab_page, self._t("tab_admin"))
            logging.info("Admin tab setup.")
            # Actual content for this tab will be in Step 11
        else:
            # Ensure admin tab is removed if user is not admin (e.g., after logout/login as non-admin)
            if hasattr(self, 'admin_tab_page') and self.admin_tab_page:
                idx = self.tab_widget.indexOf(self.admin_tab_page)
                if idx != -1:
                    self.tab_widget.removeTab(idx)
                    self.admin_tab_page.deleteLater()
                    del self.admin_tab_page
                    logging.info("Admin tab removed for non-admin user.")


    def setup_tabs_post_login(self):
        """Clears existing tabs and sets up tabs based on user role."""
        if not hasattr(self, 'tab_widget'):
            logging.error("Tab widget not initialized, cannot setup tabs.")
            return

        # Clear all existing tabs and delete their pages
        while self.tab_widget.count() > 0:
            page_widget = self.tab_widget.widget(0)
            self.tab_widget.removeTab(0)
            if page_widget:
                page_widget.deleteLater() # Important to free resources

        # Reset tab page references
        if hasattr(self, 'upload_tab_page'): del self.upload_tab_page
        if hasattr(self, 'manage_tab_page'): del self.manage_tab_page
        if hasattr(self, 'settings_tab_page'): del self.settings_tab_page
        if hasattr(self, 'admin_tab_page'): del self.admin_tab_page

        logging.info("Setting up tabs post-login...")
        self.setup_upload_tab_qt()
        self.setup_manage_tab_qt()
        self.setup_settings_tab_qt()
        self.setup_admin_tab_qt() # This will only add if user is admin


    def add_logout_button_qt(self):
        """Adds the logout button to its placeholder in the header."""
        if not hasattr(self, 'logout_button_placeholder_layout'): return # Safety check

        while self.logout_button_placeholder_layout.count():
            child = self.logout_button_placeholder_layout.takeAt(0)
            if child.widget():
                child.widget().deleteLater()

        self.logout_button = QtWidgets.QPushButton(self._t("ctkbutton_text_logout"))
        # Basic styling, can be enhanced with QPalette or stylesheets
        self.logout_button.setStyleSheet("QPushButton { background-color: #dc3545; color: white; border: 1px solid #dc3545; padding: 5px; border-radius: 3px; } QPushButton:hover { background-color: #c82333; }")
        self.logout_button.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        self.logout_button.clicked.connect(self.logout_qt)
        self.logout_button_placeholder_layout.addWidget(self.logout_button)
        self.logout_button_placeholder_widget.setVisible(True) # Ensure placeholder is visible
        logging.info("Logout button added to header.")

    def logout_qt(self):
        """Handles user logout."""
        if not self.current_user:
            return

        logging.info(f"User '{self.current_user['username']}' logging out.")

        self.current_user = None

        if hasattr(self, 'user_info_label'): self.user_info_label.setText("")
        self.setWindowTitle(self._t("app_title"))

        if hasattr(self, 'logout_button') and self.logout_button:
            self.logout_button.hide() # Hide it first
            self.logout_button.deleteLater() # Schedule for deletion
            delattr(self, 'logout_button')
        if hasattr(self, 'logout_button_placeholder_widget'): # Also hide the placeholder if desired
             self.logout_button_placeholder_widget.setVisible(False)


        if hasattr(self, 'tab_widget'):
            while self.tab_widget.count() > 0:
                widget_to_remove = self.tab_widget.widget(0)
                self.tab_widget.removeTab(0)
                if widget_to_remove:
                    widget_to_remove.deleteLater() # Important to clean up tab pages

        # Hide main window. Login dialog will be shown by run_login_dialog.
        self.hide()
        self.run_login_dialog()

    def run_login_dialog(self):
        """Placeholder for showing the login dialog."""
        # For now, assume login is successful and set a dummy user
        # This will be replaced by a proper QDialog in a later step.
        # self.current_user = {"username": "testuser", "role": "admin"}
        # logging.info(f"Simulated login for user: {self.current_user['username']}")
        # self.setup_main_application_ui() # Setup main UI after "login"
        # self.show()

        # Actual login dialog will be implemented later.
        # For now, let's just show the main window for structure verification.
        # self.setup_ui() # Call basic UI setup
        # self.show()
        # For now, let's simulate no login and keep the window hidden or minimal
        # The real login will be Step 12
        logging.info("Login dialog would run here. Main window will be shown after auth.")

        # For testing header, let's simulate a login which would then call setup_main_application_ui
        # This will be replaced by actual QDialog logic in Step 12

        # --- SIMULATED LOGIN FOR TESTING ---
        # In a real scenario, a QDialog would handle this.
        self.current_user = {"username": "testadmin", "role": "admin"}
        # self.current_user = {"username": "testuser", "role": "user"} # For testing non-admin

        self.setup_main_application_ui() # This will setup header and call setup_tabs_post_login
        self.add_logout_button_qt()      # Manually add logout button
        self.show()
        # --- END SIMULATED LOGIN ---


    def setup_main_application_ui(self):
        """Sets up the main UI components after successful login."""
        if not self.current_user:
            logging.error("Cannot setup main UI without a logged-in user.")
            return

        # Ensure base UI (header, status bar, main layout with tab_widget) is setup
        # self.setup_ui() was called from __init__ in previous step to create the containers.
        # Here we populate/update them.

        self._update_header_texts() # Update header with user info
        self.setup_tabs_post_login() # Setup the tabs

        self.setWindowTitle(f"{self._t('app_title')} - {self.current_user['username']} ({self._t('role_' + self.current_user['role'])})")
        logging.info("Main application UI setup/updated for logged-in user.")


    def closeEvent(self, event):
        """Handle application closing."""
        logging.info("Application close event triggered.")
        self.on_closing_qt() # Call the new cleanup method
        event.accept()

    # --- Non-UI Core Logic Methods (Ported from Tkinter version) ---

    def initialize_user_database(self):
        global users # Make sure we're updating the global users dict
        data_dir = get_data_dir()
        self.db_path = os.path.join(data_dir, "users.db") # Ensure users.db is in the right place
        logging.info(f"Initializing user database at: {self.db_path}")
        try:
            self.conn = sqlite3.connect(self.db_path)
            self.cursor = self.conn.cursor()
            self.cursor.execute("""
                CREATE TABLE IF NOT EXISTS users (
                    username TEXT PRIMARY KEY,
                    password TEXT NOT NULL,
                    role TEXT NOT NULL
                )
            """)
            self.conn.commit()
            self.load_users_from_db() # Populates the global `users` dictionary
            self.ensure_admin_user_db()
        except sqlite3.Error as e:
            logging.error(f"Database error during initialization: {e}", exc_info=True)
            # Show critical error to user and exit
            critical_error_box = QtWidgets.QMessageBox()
            critical_error_box.setIcon(QtWidgets.QMessageBox.Critical)
            critical_error_box.setWindowTitle("Database Error")
            critical_error_box.setText(f"A critical database error occurred:\n{e}\n\nApplication cannot continue.")
            critical_error_box.exec()
            sys.exit(1)


    def load_users_from_db(self):
        global users
        self.cursor.execute("SELECT username, password, role FROM users")
        rows = self.cursor.fetchall()
        users = {username: {"password": password, "role": role} for username, password, role in rows}
        logging.info(f"Loaded {len(users)} users from database into global 'users' dict.")
        # Update UserController if it exists and needs it, though it's initialized with `users`
        if hasattr(self, 'user_controller') and self.user_controller:
            self.user_controller.users = users
            logging.info("UserController updated with fresh user data.")


    def save_user_to_db(self, username, hashed_password, role):
        try:
            self.cursor.execute("""
                INSERT OR REPLACE INTO users(username, password, role)
                VALUES (?, ?, ?)
            """, (username, hashed_password, role))
            self.conn.commit()
            logging.info(f"User '{username}' saved to database.")
            # Ensure the global `users` dict and UserController are also updated
            users[username] = {"password": hashed_password, "role": role}
            if hasattr(self, 'user_controller'):
                self.user_controller.users = users # Keep UserController in sync
        except sqlite3.Error as e:
            logging.error(f"Failed to save user {username} to DB: {e}", exc_info=True)
            # Consider raising or notifying user

    def ensure_admin_user_db(self):
        global users
        self.cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
        count = self.cursor.fetchone()[0]
        if count == 0:
            admin_hash = pbkdf2_sha256.hash(DEFAULT_ADMIN_PASSWORD)
            self.save_user_to_db("admin", admin_hash, "admin") # This also updates global `users`
            logging.warning("No admin users found, created default admin user ('admin'/'admin123').")

    def verify_user_credentials(self, username, password):
        """Verifies username and password against the loaded user data."""
        if username in users and pbkdf2_sha256.verify(password, users[username]["password"]):
            return users[username]["role"]
        return None

    def sanitize_path(self, text):
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            text = text.replace(char, '_')
        if platform.system() == 'Windows':
            text = text.rstrip('. ')
            if len(text) > 200: # Conservative threshold for long paths
                return f"\\\\?\\{text}"
        return text

    def create_company_structure(self, company_name):
        try:
            safe_company_name = self.sanitize_path(company_name)
            self.current_company = {"display_name": company_name, "safe_name": safe_company_name}
            base_path = os.path.join(self.archives_path, safe_company_name)
            logging.info(f"Creating/Verifying structure for company: {company_name} (Safe Path: {safe_company_name})")
            os.makedirs(base_path, exist_ok=True)

            for header, sub_level_data in self.structure.items():
                header_path = os.path.join(base_path, header)
                os.makedirs(header_path, exist_ok=True)
                if isinstance(sub_level_data, dict):
                    for subheader, section_dict in sub_level_data.items():
                        subheader_path = os.path.join(header_path, subheader)
                        os.makedirs(subheader_path, exist_ok=True)
                        if isinstance(section_dict, dict):
                            for section, subsection_list in section_dict.items():
                                section_path = os.path.join(subheader_path, section)
                                os.makedirs(section_path, exist_ok=True)
                                if isinstance(subsection_list, list):
                                    for subsection in subsection_list:
                                        if subsection:
                                             subsection_path = os.path.join(section_path, subsection)
                                             os.makedirs(subsection_path, exist_ok=True)
                elif isinstance(sub_level_data, list):
                    for subheader_item in sub_level_data:
                        if subheader_item:
                             subheader_path = os.path.join(header_path, subheader_item)
                             os.makedirs(subheader_path, exist_ok=True)
            logging.debug(f"Structure verification complete for: {safe_company_name}")
        except Exception as e:
            logging.error(f"Failed to create company structure for '{company_name}': {e}", exc_info=True)
            raise # Re-raise for the caller to handle (e.g., show QMessageBox)

    def perform_file_upload(self, company_name, header, subheader, section, subsection,
                            source_file_path, intended_destination_filename):
        dest_path = "Unknown"
        safe_company_name = "Unknown"
        logging.info(f"[UploadLogic] Args: Co='{company_name}', H='{header}', S='{subheader}', Sec='{section}', SubSec='{subsection}', SrcPath='{source_file_path}', DestName='{intended_destination_filename}'")
        try:
            if not hasattr(self, 'current_company') or self.current_company.get("display_name") != company_name:
                 self.create_company_structure(company_name) # This will set self.current_company
            safe_company_name = self.current_company["safe_name"]

            structure_options = self.structure.get(header, [])
            has_subsections_defined = False
            if isinstance(structure_options, dict):
                section_dict = structure_options.get(subheader, {})
                if section and section in section_dict:
                    subsections_list = section_dict.get(section, [])
                    has_subsections_defined = bool(subsections_list)

            dest_path = os.path.join(self.archives_path, safe_company_name, header)
            if isinstance(structure_options, dict):
                if subheader: dest_path = os.path.join(dest_path, subheader)
                if section:
                    dest_path = os.path.join(dest_path, section)
                    if subsection and has_subsections_defined:
                        dest_path = os.path.join(dest_path, subsection)
            elif subheader:
                dest_path = os.path.join(dest_path, subheader)

            os.makedirs(dest_path, exist_ok=True)
            dest_file = os.path.join(dest_path, intended_destination_filename)

            if os.path.exists(dest_file):
                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                base, ext = os.path.splitext(intended_destination_filename)
                backup_name = f"{base}_backup_{timestamp}{ext}"
                backup_path = os.path.join(dest_path, backup_name)
                count = 0
                while os.path.exists(backup_path):
                     count += 1
                     backup_name = f"{base}_backup_{timestamp}_{count}{ext}"
                     backup_path = os.path.join(dest_path, backup_name)
                os.rename(dest_file, backup_path)
                logging.info(f"Existing file versioned: {dest_file} -> {backup_path}")

            shutil.copy2(source_file_path, dest_file)
            logging.info(f"File copied successfully: {source_file_path} -> {dest_file}")
            return True
        except Exception as e:
            logging.error(f"UploadLogic Error processing {source_file_path} -> {dest_path}/{intended_destination_filename}: {e}", exc_info=True)
            raise # Re-raise for caller (often a worker thread) to handle or log

    def create_archive_folder(self): # From original Tkinter version
        if not os.path.exists(self.archives_path):
            os.makedirs(self.archives_path)
            logging.info(f"Archives folder created: {self.archives_path}")
        self.secure_archive_folder() # Initial secure/hide

    def secure_archive_folder(self): # From original Tkinter version
        # This is minimal, real security is more complex
        if platform.system() == "Windows":
            try:
                subprocess.run(['attrib', '+h', '+s', self.archives_path], shell=True, check=False) # Non-recursive for the main folder
            except Exception as e:
                logging.warning(f"Could not initially hide/secure archive folder '{self.archives_path}': {e}")
        else: # Linux/macOS
            try:
                os.chmod(self.archives_path, 0o700) # rwx for owner only
            except Exception as e:
                logging.warning(f"Could not set permissions for archive folder '{self.archives_path}': {e}")
        logging.info(f"Attempted to secure archive folder: {self.archives_path}")

    def hide_archive_folder(self):
        """Hide the archive folder (non-recursively). Qt version."""
        try:
            archive_path_abs = os.path.abspath(self.archives_path)
            if not os.path.exists(archive_path_abs):
                logging.warning(f"Archive path does not exist for hiding: {archive_path_abs}")
                return

            logging.info(f"Hiding archive folder: {archive_path_abs}")
            if platform.system() == "Windows":
                subprocess.run(['attrib', '+h', '+s', archive_path_abs], shell=True, check=False)
            else:
                subprocess.run(['chmod', '700', archive_path_abs], check=False) # rwx for owner only
            logging.info("Archive folder hide command executed.")
        except Exception as e:
            logging.error(f"Error hiding archive folder: {e}", exc_info=True)

    def show_archive_folder(self):
        """Show the archive folder. Qt version."""
        try:
            archive_path_abs = os.path.abspath(self.archives_path)
            if not os.path.exists(archive_path_abs):
                logging.warning(f"Archive path does not exist for showing: {archive_path_abs}")
                return
            logging.info(f"Showing archive folder: {archive_path_abs}")
            if platform.system() == "Windows":
                subprocess.run(['attrib', '-h', '-s', archive_path_abs], shell=True, check=False)
            else:
                subprocess.run(['chmod', '755', archive_path_abs], check=False) # rwxr-xr-x
            logging.info("Archive folder show command executed.")
        except Exception as e:
            logging.error(f"Error showing archive folder: {e}", exc_info=True)


    def on_closing_qt(self):
        """Handle application closing for PySide6."""
        logging.info("on_closing_qt called.")
        if self.current_user:
            logging.info(f"User '{self.current_user['username']}' logging out due to app close.")
            # self.logout_logic() # If there's specific logout logic beyond clearing current_user

        # Stop Watchdog observer if it's running
        # This will be integrated in Step 13
        # if hasattr(self, 'observer') and self.observer.is_alive():
        #     self.observer.stop()
        #     self.observer.join(timeout=1) # Wait briefly for it to stop
        #     logging.info("Watchdog observer stopped.")

        # Hide archive folder
        self.hide_archive_folder()

        # Wait for all threads in QThreadPool to finish
        # This is important if there are background tasks like file operations
        if hasattr(self, 'thread_pool'):
            logging.info("Waiting for QThreadPool to finish...")
            self.thread_pool.waitForDone(-1) # Wait indefinitely
            logging.info("QThreadPool finished.")

        # Close database connection
        if hasattr(self, 'conn') and self.conn:
            self.conn.close()
            logging.info("Database connection closed.")

        logging.info("Application cleanup complete. Exiting.")


if __name__ == "__main__":
    # Ensure QApplication is created before any QObjects, including FileArchiveApp
    if not QtWidgets.QApplication.instance():
        app = QtWidgets.QApplication(sys.argv)
    else:
        app = QtWidgets.QApplication.instance()

    # Set application name and version (optional, but good for some OS features)
    app.setApplicationName("FileArchiveApp")
    app.setApplicationVersion("1.0.Ps") # Ps for PySide

    # Create and show the main application window
    main_app_window = FileArchiveApp()

    # The original Tkinter app showed a splash screen then login.
    # For PySide6, login will be a QDialog.
    # We'll call a method to handle login flow.
    # For now, to test structure, we can bypass login or show main window directly.
    # main_app_window.show() # Show main window directly for now.
    # Login dialog will be shown from main_app_window.run_login_dialog() or similar
    # For now, just show the window to verify structure and non-UI logic init.
    # The UI elements will be built progressively.

    # Instead of showing directly, the login flow will handle visibility.
    # main_app_window.setup_ui() # Basic UI elements for testing
    # main_app_window.show()

    # The actual display of the window will be handled after the login dialog (Step 12)
    # For now, we are just setting up the class.
class FileArchiveApp(QtWidgets.QMainWindow): # Ensure this is below Worker definition
    # ... (existing FileArchiveApp __init__ and other methods) ...

    # --- Slots for Worker Signals ---
    def _handle_worker_result(self, task_name, result_data):
        """Handles results from a specific worker task."""
        logging.info(f"Worker task '{task_name}' completed with result: {result_data}")
        # Example: self.statusBar().showMessage(f"Task {task_name} successful.", 5000)

    def _handle_worker_error(self, task_name, error_message):
        """Handles errors from a specific worker task."""
        logging.error(f"Worker task '{task_name}' error: {error_message}")
        QtWidgets.QMessageBox.critical(self,
                                   self._t("error_title_task_failed", task=task_name),
                                   f"{self._t('error_background_task_general', task=task_name)}\n\n{error_message}")
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_task_failed_with_name", task=task_name), 5000)

    def _handle_worker_finished(self, task_name):
        """Handles completion of a specific worker task (success or failure)."""
        logging.info(f"Worker task '{task_name}' finished.")
        if hasattr(self, 'statusBar'):
            current_message = self.statusBar().currentMessage()
            if not current_message or "failed" not in current_message.lower() and "error" not in current_message.lower():
                 self.statusBar().showMessage(self._t("status_task_completed_ready", task=task_name), 3000)
            # Reset progress bar if it was used for this task
            if hasattr(self, 'progress_bar') and self.progress_bar.isVisible():
                self.progress_bar.setValue(0)
                self.progress_bar.setVisible(False)


    # --- Refactored Methods Using Worker ---

    def refresh_folders_qt(self):
        """Refreshes folder visibility (shows all). Uses Qt Worker."""
        task_name = "RefreshFolders"
        logging.info(f"Admin triggered {task_name} (Qt Worker).")

        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_refreshing_folders"), 0)
        # Placeholder: self.admin_system_tab.refresh_button.setEnabled(False)

        worker = Worker(self._task_refresh_folders)
        # Connect signals with task name for context
        worker.result.connect(lambda result_data: self._handle_refresh_folders_result(task_name, result_data))
        worker.error.connect(lambda err_msg: self._handle_worker_error(task_name, err_msg))
        worker.finished.connect(lambda: self._handle_worker_finished(task_name))
        worker.finished.connect(lambda: setattr(self, f'{task_name}_worker', None)) # Clean up worker ref

        setattr(self, f'{task_name}_worker', worker) # Store worker reference if needed
        self.thread_pool.start(worker)

    def _task_refresh_folders(self):
        """The actual logic for refreshing folders (to be run in worker)."""
        try:
            if hasattr(self, 'archive_controller'):
                self.archive_controller.clear_cache()
            for root, dirs, _ in os.walk(self.archives_path):
                for d in dirs:
                    show_path = os.path.join(root, d)
                    if platform.system() == "Windows":
                        subprocess.run(['attrib', '-h', '-s', show_path], shell=True, check=False)
            self.show_archive_folder()
            return self._t("status_folders_refreshed_visible")
        except Exception as e:
            logging.error(f"Error in _task_refresh_folders: {e}", exc_info=True)
            raise

    def _handle_refresh_folders_result(self, task_name, message):
        QtWidgets.QMessageBox.information(self, self._t("info_title_refreshed"), message)
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(message, 5000)
        # Placeholder: self.admin_system_tab.refresh_button.setEnabled(True)


    def search_archive_qt(self, query, file_type_filter, start_date_str, end_date_str, search_dialog_ref):
        task_name = "ArchiveSearch"
        logging.info(f"{task_name} initiated: Q='{query}', Type='{file_type_filter}', Start='{start_date_str}', End='{end_date_str}'")
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_searching_archive"), 0)
        if search_dialog_ref and hasattr(search_dialog_ref, 'search_button'):
            search_dialog_ref.search_button.setEnabled(False)
            search_dialog_ref.status_label.setText(self._t("status_searching_progress"))


        worker = Worker(self._task_perform_search, query, file_type_filter, start_date_str, end_date_str)
        worker.result.connect(lambda results: self._handle_search_results(task_name, results, search_dialog_ref))
        worker.error.connect(lambda err_msg: self._handle_search_error(task_name, err_msg, search_dialog_ref))
        worker.finished.connect(lambda: self._handle_worker_finished(task_name))
        worker.finished.connect(lambda: setattr(self, f'{task_name}_worker', None))

        setattr(self, f'{task_name}_worker', worker)
        self.thread_pool.start(worker)

    def _task_perform_search(self, query, file_type_filter, start_date_str, end_date_str):
        # (Search logic as previously defined, ensure it raises exceptions on error)
        # ... (omitted for brevity, it's the same core logic as before)
        # For demonstration, assume it returns a list of result dicts or raises error
        query_lower = query.strip().lower()
        if query_lower:
             self.search_queries.append(query_lower) # QMutex might be needed if list is complexly managed

        try:
            start_date = datetime.datetime.strptime(start_date_str, "%Y-%m-%d") if start_date_str else None
            end_date = datetime.datetime.strptime(end_date_str, "%Y-%m-%d") if end_date_str else None
        except ValueError as ve:
            raise ValueError(self._t("error_invalid_date_format_yyyymmdd") + f": {ve}")

        results = []
        # (Walk filesystem and build results list - same as before)
        for root, dirs, files_in_dir in os.walk(self.archives_path):
            for name in dirs + files_in_dir:
                if query_lower in name.lower():
                    full_path = os.path.join(root, name)
                    is_f = os.path.isfile(full_path)
                    if is_f:
                        if file_type_filter != "All":
                            ext = os.path.splitext(name)[1].lower()
                            if file_type_filter == "Images" and ext not in IMAGE_EXTENSIONS: continue
                            elif file_type_filter == "Documents" and ext not in DOCUMENT_EXTENSIONS: continue
                    try: mod_time = datetime.datetime.fromtimestamp(os.path.getmtime(full_path))
                    except: mod_time = None
                    if start_date and mod_time and mod_time < start_date: continue
                    if end_date and mod_time and mod_time > end_date: continue
                    results.append({"path": full_path, "name": name, "is_dir": not is_f,
                                    "mod_time": mod_time.strftime("%Y-%m-%d %H:%M:%S") if mod_time else "N/A",
                                    "size": os.path.getsize(full_path) if is_f else "N/A"})
        return results


    def _handle_search_results(self, task_name, results, search_dialog_ref):
        logging.info(f"{task_name} results received: {len(results)} items.")
        if search_dialog_ref:
            search_dialog_ref.display_results(results) # Assume dialog has this method
            search_dialog_ref.search_button.setEnabled(True)
            search_dialog_ref.status_label.setText(self._t("status_search_complete_results_short", count=len(results)))
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_search_complete_results", count=len(results)), 5000)

    def _handle_search_error(self, task_name, error_message, search_dialog_ref):
        logging.error(f"{task_name} error: {error_message}")
        QtWidgets.QMessageBox.critical(search_dialog_ref if search_dialog_ref else self,
                                   self._t("search_error_title"), error_message)
        if search_dialog_ref:
            search_dialog_ref.search_button.setEnabled(True)
            search_dialog_ref.status_label.setText(self._t("search_failed_status"))
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_search_failed"), 5000)

    def add_structure_element_qt(self, company_display_name, header, current_subheader, current_section, add_type, new_name_sanitized, dialog_ref):
        task_name = "AddFolder"
        logging.info(f"Initiating {task_name}: Type='{add_type}', Name='{new_name_sanitized}'")
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_adding_folder", name=new_name_sanitized),0)
        if dialog_ref and hasattr(dialog_ref, 'add_button'):
            dialog_ref.add_button.setEnabled(False)

        worker = Worker(self._task_add_folder, company_display_name, header, current_subheader, current_section, add_type, new_name_sanitized)
        worker.result.connect(lambda res_info: self._handle_add_folder_result(task_name, res_info, dialog_ref))
        worker.error.connect(lambda err_msg: self._handle_add_folder_error(task_name, err_msg, dialog_ref))
        worker.finished.connect(lambda: self._handle_worker_finished(task_name))
        worker.finished.connect(lambda: setattr(self, f'{task_name}_worker', None))

        setattr(self, f'{task_name}_worker', worker)
        self.thread_pool.start(worker)

    def _task_add_folder(self, company_display_name, header, current_subheader, current_section, add_type, new_name):
        # (Logic from previous _task_add_folder, ensure it returns a dict with "success": True/False and other info)
        # ... (omitted for brevity, same core logic as before)
        target_level_description = "Unknown"
        parent_path_of_new_folder = None
        try:
            safe_company_name = self.sanitize_path(company_display_name)
            if not hasattr(self, 'current_company') or self.current_company.get("safe_name") != safe_company_name:
                self.create_company_structure(company_display_name)
                safe_company_name = self.current_company["safe_name"]

            parent_path = os.path.join(self.archives_path, safe_company_name, header)
            if add_type == "Subheader": parent_path_of_new_folder = parent_path; target_level_description = f"Header '{header}'"
            elif add_type == "Section":
                if not current_subheader: raise ValueError("Subheader missing for Section.")
                parent_path = os.path.join(parent_path, current_subheader); parent_path_of_new_folder = parent_path; target_level_description = f"Subheader '{current_subheader}'"
            elif add_type == "Subsection":
                if not current_subheader or not current_section: raise ValueError("Context missing for Subsection.")
                parent_path = os.path.join(parent_path, current_subheader, current_section); parent_path_of_new_folder = parent_path; target_level_description = f"Section '{current_section}'"
            else: raise ValueError(f"Invalid add_type: {add_type}")

            new_element_path = os.path.join(parent_path, new_name)
            os.makedirs(new_element_path, exist_ok=True)
            user = self.current_user['username'] if self.current_user else 'System'
            logging.info(f"User '{user}' created folder: '{new_element_path}'")
            return {"success": True, "new_name": new_name, "target_level_description": target_level_description,
                    "parent_path_of_new_folder": parent_path_of_new_folder}
        except Exception as e:
            return {"success": False, "error_message": str(e), "new_name": new_name}


    def _handle_add_folder_result(self, task_name, result_info, dialog_ref):
        if dialog_ref and hasattr(dialog_ref, 'add_button'):
            dialog_ref.add_button.setEnabled(True)
        if hasattr(self, 'statusBar'): self.statusBar().clearMessage()

        if result_info.get("success"):
            new_name = result_info["new_name"]; target_desc = result_info["target_level_description"]
            parent_path = result_info.get("parent_path_of_new_folder")
            QtWidgets.QMessageBox.information(dialog_ref or self, self._t("info_title_folder_created"),
                                         self._t("status_folder_created_success", name=new_name, context=target_desc))
            if parent_path and hasattr(self, 'archive_controller'):
                self.archive_controller.clear_cache(path_prefix=parent_path)
            if hasattr(self, 'update_options_qt'): self.update_options_qt() # Refresh ComboBoxes
            if dialog_ref: dialog_ref.accept()
        else:
            QtWidgets.QMessageBox.critical(dialog_ref or self, self._t("error_title_folder_creation"),
                                      self._t("error_folder_creation_failed", name=result_info.get("new_name","?"), error=result_info.get("error_message","?")))
        if hasattr(self, 'statusBar'): self.statusBar().showMessage(self._t("ctklabel_text_ready"), 3000)

    def _handle_add_folder_error(self, task_name, err_msg, dialog_ref): # Matches general error handler
        self._handle_worker_error(task_name, err_msg) # Use general handler
        if dialog_ref and hasattr(dialog_ref, 'add_button'):
            dialog_ref.add_button.setEnabled(True)


    def batch_upload_qt(self, company_name, header, subheader, section, subsection, file_paths, auto_rename_confirmed, calling_dialog_ref=None):
        task_name = "BatchUpload"
        logging.info(f"{task_name} initiated for {len(file_paths)} files.")
        if hasattr(self, 'statusBar'):
            self.statusBar().showMessage(self._t("status_batch_upload_starting", count=len(file_paths)), 0)
        if hasattr(self, 'progress_bar'):
            self.progress_bar.setValue(0); self.progress_bar.setVisible(True)

        # Example: Disable the button in the calling dialog/tab if reference is passed
        if calling_dialog_ref and hasattr(calling_dialog_ref, 'batch_upload_button'):
            calling_dialog_ref.batch_upload_button.setEnabled(False)

        worker = Worker(self._task_batch_upload, company_name, header, subheader, section, subsection, file_paths, auto_rename_confirmed)
        worker.result.connect(lambda res_info: self._handle_batch_upload_result(task_name, res_info, calling_dialog_ref))
        worker.error.connect(lambda err_msg: self._handle_batch_upload_error(task_name, err_msg, calling_dialog_ref))
        worker.progress.connect(self._update_batch_progress)
        worker.finished.connect(lambda: self._handle_worker_finished(task_name))
        worker.finished.connect(lambda: setattr(self, f'{task_name}_worker', None))
        # Re-enable button on finish
        if calling_dialog_ref and hasattr(calling_dialog_ref, 'batch_upload_button'):
             worker.finished.connect(lambda: calling_dialog_ref.batch_upload_button.setEnabled(True))


        setattr(self, f'{task_name}_worker', worker)
        self.thread_pool.start(worker)

    def _task_batch_upload(self, company_name, header, subheader, section, subsection, file_paths, auto_rename_confirmed):
        # (Logic from previous _task_batch_upload - ensure it emits progress via self.progress.emit(percentage))
        # ... (omitted for brevity, same core logic, ensure self.progress.emit() is called inside the loop)
        total_files = len(file_paths); success_count = 0; naming_failures = []; other_errors = []
        try: self.create_company_structure(company_name)
        except Exception as e_struct: raise Exception(f"Batch structure error: {e_struct}")

        for i, fp in enumerate(file_paths):
            original_filename = os.path.basename(fp)
            intended_filename = original_filename
            try:
                required_prefix = "" # Calculate required_prefix as before
                structure_options_task = self.structure.get(header, [])
                if isinstance(structure_options_task, dict):
                    section_dict_task = structure_options_task.get(subheader, {})
                    if section and section in section_dict_task:
                        subsections_list_task = section_dict_task.get(section, [])
                        if subsection and subsection in subsections_list_task: required_prefix = subsection
                        else: required_prefix = section
                    elif subheader: required_prefix = subheader
                elif subheader: required_prefix = subheader

                needs_rename = False
                if required_prefix and not (original_filename.startswith(required_prefix + "_") or original_filename.startswith(required_prefix)):
                    needs_rename = True

                if needs_rename:
                    if auto_rename_confirmed: intended_filename = f"{required_prefix}_{original_filename}"
                    else: naming_failures.append(original_filename); self.progress.emit(int(((i + 1) / total_files) * 100)); continue

                self.perform_file_upload(company_name, header, subheader, section, subsection, fp, intended_filename)
                success_count += 1
            except Exception as e_file: other_errors.append(f"{original_filename}: {e_file}")
            self.progress.emit(int(((i + 1) / total_files) * 100))
        return {"total": total_files, "success": success_count, "naming_fails": naming_failures, "other_errs": other_errors}


    def _update_batch_progress(self, percentage):
        if hasattr(self, 'progress_bar'): self.progress_bar.setValue(percentage)
        if hasattr(self, 'statusBar'): self.statusBar().showMessage(self._t("status_batch_progress", progress=percentage),0)

    def _handle_batch_upload_result(self, task_name, result_info, calling_dialog_ref):
        self.report_batch_results_qt(result_info["total"], result_info["success"], result_info["naming_fails"], result_info["other_errs"])
        if hasattr(self, 'progress_bar'): self.progress_bar.setVisible(False)
        if hasattr(self, 'statusBar'): self.statusBar().showMessage(self._t("status_batch_upload_completed"), 5000)
        if calling_dialog_ref and hasattr(calling_dialog_ref, 'batch_upload_button'): # Re-enable button if passed
             calling_dialog_ref.batch_upload_button.setEnabled(True)


    def _handle_batch_upload_error(self, task_name, error_message, calling_dialog_ref):
        self._handle_worker_error(task_name, error_message) # Use general handler
        if hasattr(self, 'progress_bar'): self.progress_bar.setVisible(False)
        if calling_dialog_ref and hasattr(calling_dialog_ref, 'batch_upload_button'): # Re-enable button
             calling_dialog_ref.batch_upload_button.setEnabled(True)


    def report_batch_results_qt(self, total, success, naming_fails, other_errs):
        # (Logic from previous report_batch_results_qt)
        # ... (omitted for brevity, same core logic)
        message_lines = [self._t("batch_report_header", success=success, total=total)]
        status_text = self._t("batch_status_completed", success=success, total=total)
        if naming_fails:
            message_lines.append(f"\n{self._t('batch_report_naming_skipped_header', count=len(naming_fails))}")
            message_lines.extend([f"- {nf}" for nf in naming_fails])
            status_text += f" {self._t('batch_status_naming_skipped', count=len(naming_fails))}"
        if other_errs:
            message_lines.append(f"\n{self._t('batch_report_other_errors_header', count=len(other_errs))}")
            message_lines.extend([f"- {oe}" for oe in other_errs]) # This line was correct
            status_text += f" {self._t('batch_status_other_errors', count=len(other_errs))}"

        detailed_message = "\n".join(message_lines)
        if len(detailed_message) > 1500: detailed_message = detailed_message[:1500] + "\n\n..." + self._t("batch_report_see_log_details")

        if naming_fails or other_errs: QtWidgets.QMessageBox.warning(self, self._t("batch_report_title_issues"), detailed_message)
        else: QtWidgets.QMessageBox.information(self, self._t("batch_report_title_complete"), detailed_message)

        if hasattr(self, 'statusBar'): self.statusBar().showMessage(status_text, 7000)
        if hasattr(self, 'progress_bar'):
            QtCore.QTimer.singleShot(3000, lambda: (self.progress_bar.setValue(0), self.progress_bar.setVisible(False)))


if __name__ == "__main__":
