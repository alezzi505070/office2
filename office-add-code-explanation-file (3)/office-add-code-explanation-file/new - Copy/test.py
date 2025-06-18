import sys
import os

if getattr(sys, 'frozen', False):
    base_dir = sys._MEIPASS  # PyInstaller's temporary folder
    # Set the Tcl/Tk library paths based on our bundled structure
    os.environ["TCL_LIBRARY"] = os.path.join(base_dir, "tcl", "tcl8.6")
    os.environ["TK_LIBRARY"] = os.path.join(base_dir, "tcl", "tk8.6")


import sqlite3
import re
import tempfile
import shutil
import platform
import subprocess
import logging
import datetime
import threading
import queue

import win32print
# Add this near your other imports
# Import the module itself, and specific functions you need.
# DO NOT import CURRENT_LANGUAGE directly.
from translations import load_translations, set_language, get_translation
import translations # <-- Add this line

from concurrent.futures import ThreadPoolExecutor
import json
import customtkinter as ctk
from PIL import Image, ImageTk
from tkinter import filedialog, messagebox, simpledialog
import tkinter as tk
# For secure password hashing (using passlib)
from passlib.hash import pbkdf2_sha256
import time
# Controllers for MVC pattern
from controllers.user_controller import UserController
from controllers.archive_controller import ArchiveController
from concurrent.futures import ThreadPoolExecutor
import cProfile
import pstats
import threading
# For real-time monitoring using watchdog
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
WATCHDOG_AVAILABLE = True
# WIA import assumed to be handled elsewhere or within try block
try:
    import comtypes.client
    WIA_AVAILABLE = True
except ImportError:
    WIA_AVAILABLE = False
    logging.warning("comtypes.client not found. Scanning via WIA will be disabled.")

# For drag and drop functionality
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DRAG_DROP_ENABLED = True
except ImportError:
    DRAG_DROP_ENABLED = False
    print("TkinterDnD not installed. Drag and drop will be disabled.")
    print("Install with: pip install tkinterdnd2")

# ------------------------------------------------------------------------------
# Constants (avoid magic strings)
# ------------------------------------------------------------------------------
DEFAULT_ADMIN_PASSWORD = "admin123"
DEFAULT_USER_PASSWORD = "user123"
IMAGE_EXTENSIONS = [".jpg", ".jpeg", ".png", ".gif", ".bmp", ".tiff"]
DOCUMENT_EXTENSIONS = [".xlsx", ".xls", ".doc", ".docx", ".ppt", ".pptx", ".pdf"] # Added document extensions
SUPPORTED_FILE_EXTENSIONS = IMAGE_EXTENSIONS + DOCUMENT_EXTENSIONS # Combined list
set_language("en") # Or "ar" if you want Arabic default

# ------------------------------------------------------------------------------
# Logging Configuration
# ------------------------------------------------------------------------------
# ------------------------------------------------------------------------------
# Logging Configuration (CORRECTED FOR PACKAGED APPS)
# ------------------------------------------------------------------------------
# --- Custom Handler for Live Log Writing ---
class LiveWritingFileHandler(logging.FileHandler):
    """
    A custom logging file handler that flushes the buffer after every record.
    This ensures that log entries are written to the file in real-time.
    """
    def emit(self, record):
        # Call the original emit method to format and write the record
        super().emit(record)
        # Immediately flush the stream to disk
        self.flush()

# ------------------------------------------------------------------------------
# Logging Configuration (MODIFIED FOR LIVE WRITING)
# ------------------------------------------------------------------------------
def get_data_dir():
    """
    Returns the correct directory for data files (logs, db)
    based on whether the app is a script or a frozen .exe.
    """
    if getattr(sys, 'frozen', False):
        # Running as a bundled .exe
        # Use the standard AppData folder for persistent data
        return os.path.join(os.getenv('APPDATA'), 'FileArchiveApp')
    else:
        # Running as a .py script
        # Use the script's directory for local development data
        return os.path.dirname(os.path.abspath(__file__))

def setup_logging():
    """Configures logging to a location appropriate for script or exe."""
    # --- THIS IS THE KEY CHANGE ---
    data_dir = get_data_dir()
    # ----------------------------

    # Create the directory if it doesn't exist
    os.makedirs(data_dir, exist_ok=True)

    log_file_path = os.path.join(data_dir, 'archive_app.log')
    

    # Configure logging
    log_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler = LiveWritingFileHandler(log_file_path)
    handler.setFormatter(log_formatter)
    logger = logging.getLogger()
    logger.setLevel(logging.INFO) # Set your desired level here

    # Clear existing handlers to prevent duplicate logs if this function is called again
    if logger.hasHandlers():
        logger.handlers.clear()

    logger.addHandler(handler)

    # Also add a handler to print to console (very useful for debugging)
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(log_formatter)
    logger.addHandler(console_handler)

    logging.info(f"--- Logging initialized (Live Writing). Log file at: {log_file_path} ---")

setup_logging()
logging.critical("CRITICAL_LOG: Logging configured at DEBUG level.") # Added this line
# ------------------------------------------------------------------------------
# In-Memory User Store (for demo purposes only; use a database or external config in production)
# ------------------------------------------------------------------------------
users = {
    "admin": {"password": pbkdf2_sha256.hash(DEFAULT_ADMIN_PASSWORD), "role": "admin"},
    "user": {"password": pbkdf2_sha256.hash(DEFAULT_USER_PASSWORD), "role": "user"}
}
# Python
def choose_printer():
    printers = [printer[2] for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)]
    if not printers:
        return None
    # Use a customtkinter Toplevel to match other windows
    top = ctk.CTkToplevel()
    top.title("Select Printer")
    top.geometry("300x120")
    top.grab_set()        # Enforce modality
    top.focus_force()     # Bring the window to the front
    selected_printer = tk.StringVar(value=printers[0])
    ctk.CTkLabel(top, text=get_translation("ctklabel_text_select_a_printer"), font=("Segoe UI", 14)).pack(pady=10)
    option_menu = ctk.CTkOptionMenu(top, variable=selected_printer, values=printers, font=("Segoe UI", 14))
    option_menu.pack(padx=20)
    ctk.CTkButton(top, text=get_translation("ctkbutton_text_ok"), command=top.destroy, font=("Segoe UI", 14)).pack(pady=10)
    top.wait_window()
    return selected_printer.get()

# ------------------------------------------------------------------------------
# Watchdog Event Handler for Real-Time Monitoring
# ------------------------------------------------------------------------------
class ArchiveEventHandler(FileSystemEventHandler):
    def __init__(self, ui_queue, notification_callback):
        super().__init__()
        self.ui_queue = ui_queue
        self.notification_callback = notification_callback
        self.last_event = 0

    def on_any_event(self, event):
        now = time.time()
        # Throttle: update at most every 1.5 seconds
        if now - self.last_event < 1.5:
            return
        self.last_event = now
        update_text = f"Archive updated at {datetime.datetime.now().strftime('%H:%M:%S')}"
        self.ui_queue.put(lambda: self.notification_callback(update_text))


# ------------------------------------------------------------------------------
# File Archive Application with Modern UI Improvements
# ------------------------------------------------------------------------------
class FileArchiveApp:
    # ... existing code ...
    # Define valid image file extensions globally
    IMAGE_FILE_EXTENSIONS = IMAGE_EXTENSIONS # Use the constant
    SUPPORTED_FILE_EXTENSIONS = SUPPORTED_FILE_EXTENSIONS # Use the constant
    def _t(self, key):
        """Looks up the translation for a key in the dictionary."""
        # current_translations = self.translations if isinstance(self.translations, dict) else {}
        # return current_translations.get(key, f"_{key}_")
        return get_translation(key)

    def __init__(self):
        set_language("en") # This is fine as is
        logging.critical("CRITICAL_LOG: FileArchiveApp __init__ started.") # Added this line
        # Initialize current_user first, before it's referenced
        global DRAG_DROP_ENABLED
        self.current_user = None
        # self.translations = {} # Initialize translation dictionary # Removed
        # self.load_app_translations() # Load translations based on CURRENT_LANGUAGE # Removed

        # --- Database and Controllers ---
        self.initialize_user_database() # Connects and loads users into global `users` dict
        # Ensure users dict is loaded before UserController initialization if it relies on it immediately
        if not users and hasattr(self, 'load_users_from_db'): # Check if users is empty and loader exists
            self.load_users_from_db()

        self.user_controller = UserController(users) # Uses the global `users` dict

        # --- Paths and State ---
        self.archives_path = "archives"
        # Ensure base path exists early, handle potential errors during creation
        try:
             os.makedirs(self.archives_path, exist_ok=True)
        except OSError as e:
             logging.error(f"CRITICAL: Could not create archives directory '{self.archives_path}'. Error: {e}", exc_info=True)
             # You might want to show an error message and exit here if the archive path is essential
             messagebox.showerror("Fatal Error", f"Could not create required directory:\n{self.archives_path}\n\n{e}\n\nApplication cannot continue.")
             sys.exit(1) # Exit if the archive dir can't be created

        self.search_queries = []
        # self.file_comments = {} # Removed comments functionality

        # --- Threading and UI Sync ---
        self.search_queries_lock = threading.Lock()
        # self.file_comments_lock = threading.Lock() # Removed
        self.ui_queue = queue.Queue()
        # Use context manager for ThreadPoolExecutor if Python version supports it well,
        # otherwise ensure shutdown in on_closing
        self.executor = ThreadPoolExecutor(max_workers=min(4, os.cpu_count() or 1)) # Limit workers slightly
        logging.info(f"ThreadPoolExecutor created with max_workers={self.executor._max_workers}")


        # --- Document Structure Definition ---
        # Using stable English keys internally is often less complex.
        self.structure = {
            "Permanent Audit File": { # English key
                "c1": {}, "c2": {}, "c3": {}, "c4": {}, "c5": {}, "c6": {},
            },
            "Working Papers File": { # English key
                "A": {str(i): [] for i in range(1, 16)},
                "B1": {
                    "B10": ["B10A"],
                    "B11": ["B11A"],
                    "B12": [], "B13": [], "B14": [], "B15": [],
                    "B16": [], "B17": [], "B18": [], "B19": [],
                },
                "B2": {
                    "B20": [], "B21": [], "B22": [], "B23": [],
                    "B24": [], "B25": [], "B26": [], "B27": [],
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
                 # Add other sections like I3, I4, I5, I6, I7 etc. if needed
                 # "I3": { ... },
            }
        }
        # Create ArchiveController *after* self.structure is defined
        self.archive_controller = ArchiveController(self.structure, self.archives_path)


        # --- UI Setup ---
        # Set appearance mode and theme (consider loading from a settings file)
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # --- Main Window Setup (Handles Drag & Drop presence) ---
        if DRAG_DROP_ENABLED:
            try:
                self.dnd_root = TkinterDnD.Tk()
                self.dnd_root.withdraw() # Hide the dummy root
                self.dnd_root.title("DnD Handler") # For task manager clarity
                self.main_app = ctk.CTkToplevel(self.dnd_root) # App is a Toplevel
                # Ensure closing the main app also closes the dnd root
                self.main_app.protocol("WM_DELETE_WINDOW", self.on_closing)
                logging.info("TkinterDnD root created, main app is CTkToplevel.")
            except Exception as e:
                logging.error(f"Failed to initialize TkinterDnD: {e}. Disabling drag & drop.", exc_info=True)
                DRAG_DROP_ENABLED = False
                # Fallback to standard CTk window
                self.main_app = ctk.CTk()
                self.main_app.protocol("WM_DELETE_WINDOW", self.on_closing)
        else:
            self.main_app = ctk.CTk()
            self.main_app.protocol("WM_DELETE_WINDOW", self.on_closing)
            logging.info("Using standard CTk window (TkinterDnD not enabled/failed).")

        self.main_app.title(self._t("app_title")) # Use translation helper
        self.main_app.geometry("950x700") # Adjust size as needed

        # --- Main Frame (Takes up the whole window) ---
        self.main_frame = ctk.CTkFrame(self.main_app, corner_radius=0) # Use 0 radius for seamless look
        self.main_frame.pack(fill="both", expand=True, padx=0, pady=0) # Fill entire window

        self.main_frame.grid_rowconfigure(1, weight=1)  # Tab content expands vertically
        self.main_frame.grid_columnconfigure(0, weight=1) # Tabs expand horizontally

        # --- Header Frame ---
        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent", height=60) # Set consistent height
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 5)) # Padding for content inside header
        self.header_frame.grid_columnconfigure(0, weight=1) # Title expands
        self.header_frame.grid_columnconfigure(1, weight=0) # User label fixed size
        self.header_frame.grid_columnconfigure(2, weight=0) # Language button fixed size
        self.header_frame.grid_columnconfigure(3, weight=0) # Logout button fixed size <<< CONFIGURED COLUMN 3

        # App Title
        ctk.CTkLabel(self.header_frame, text=self._t("ctklabel_text_file_archiving_system"), # Use self._t
            font=("Segoe UI", 26, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 10))

        # User info label (text updated on login)
        self.user_label = ctk.CTkLabel(self.header_frame, text="", font=("Segoe UI", 12))
        self.user_label.grid(row=0, column=1, padx=10, sticky="e") # Align right before language button

        # Language switch button
        self.language_button = ctk.CTkButton(self.header_frame, text=self._t("button_text_switch_language"), # Use self._t
                                            command=self.switch_language, font=("Segoe UI", 12), width=120)
        self.language_button.grid(row=0, column=2, padx=10, sticky="e") # Align right before logout

        # Logout button placeholder (button added later by add_logout_button)
        self.logout_button_placeholder = ctk.CTkFrame(self.header_frame, fg_color="transparent", width=110) # Give it enough space
        self.logout_button_placeholder.grid(row=0, column=3, padx=(5, 0), sticky="e") # Grid in column 3


        # --- Tab View ---
        self.tabview = ctk.CTkTabview(self.main_frame, corner_radius=8)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10)) # Reduced top padding

        # Create base tabs using translation keys via helper method
        self._create_tabs() # This method now creates tab_upload, tab_manage, tab_settings

        # --- Status Bar ---
        self.status_frame = ctk.CTkFrame(self.main_frame, height=30, corner_radius=0)
        self.status_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0) # No outer padding
        self.status_frame.grid_columnconfigure(0, weight=1) # Label expands

        # Notification Label
        self.notification_label = ctk.CTkLabel(self.status_frame, text=self._t("ctklabel_text_ready"), font=("Segoe UI", 12)) # Use self._t
        self.notification_label.grid(row=0, column=0, sticky="w", padx=10) # Inner padding for text

        # Progress Bar
        self.progress_bar = ctk.CTkProgressBar(self.status_frame, width=150)
        self.progress_bar.grid(row=0, column=1, padx=10, sticky="e") # Inner padding
        self.progress_bar.set(0)

        # --- Admin Controls Flag ---
        self.admin_controls_added = False # Still used to track if admin tab was ever added

        # --- Initialization Steps ---
        # Keyboard shortcuts
        self.main_app.bind("<Control-t>", lambda event: self.toggle_theme())

        # Create archive folder and secure it (hide initially)
        self.create_archive_folder()

        # Start real-time monitoring using watchdog (if available)
        if WATCHDOG_AVAILABLE:
            self.start_monitoring()
        else:
            logging.warning("Watchdog not available, file monitoring disabled.")
            self.notification_label.configure(text=self._t("status_monitoring_disabled")) # Inform user

        # Start UI queue processing
        self.main_app.after(100, self.process_ui_queue)

        # Hide main app initially - will be shown after successful login
        self.main_app.withdraw()

        logging.info("FileArchiveApp __init__ completed.")

    # Add these methods inside your FileArchiveApp class (e.g., after the process_ui_queue method)

    def heavy_task(self, total_steps=10):
        from cython_heavy import cython_heavy_task
        progress_values = cython_heavy_task(total_steps)
        for p in progress_values:
            self.ui_queue.put(lambda p=p: self.progress_bar.set(p))
        self.ui_queue.put(lambda: self.notification_label.configure(text=get_translation("configure_text_heavy_task_complete")))

    def start_heavy_task(self):
        # Use the shared executor instead of creating a new one each time.
        self.executor.submit(self.heavy_task)

    # Add these methods inside the FileArchiveApp class
 # =========================================================================
    # ==== LANGUAGE SWITCHING (CORRECTED) =====================================
    # =========================================================================
    def switch_language(self):
        """Switches the application language and rebuilds the UI."""
        
        # 1. Determine the new language based on the current one
        new_lang = "ar" if translations.CURRENT_LANGUAGE == "en" else "en"
        logging.info(f"Attempting to switch language from '{translations.CURRENT_LANGUAGE}' to '{new_lang}'")

        # 2. Get the key of the currently selected tab so we can re-select it later
        current_display = self.tabview.get()
        selected_key = None
        
        possible_tab_keys = ["tab_upload_files", "tab_manage_files", "tab_settings"]
        if self.current_user and self.current_user["role"] == "admin" and hasattr(self, 'tab_admin') and self.tab_admin.winfo_exists():
            possible_tab_keys.append("tab_admin")

        for key in possible_tab_keys:
            # Use the current language to find the matching key
            if self._t(key) == current_display:
                selected_key = key
                logging.info(f"Identified active tab key: '{selected_key}' for displayed name '{current_display}'")
                break
        
        if not selected_key:
            logging.warning(f"Could not determine key for tab '{current_display}'. Will default to first tab.")

        # 3. Set the new language globally
        switch_success = set_language(new_lang)

        # 4. If the language was set successfully, rebuild the UI
        if switch_success:
            self._rebuild_ui_for_language(selected_key)
        else:
            messagebox.showerror("Language Error", f"Could not switch to language '{new_lang}'.")
            logging.error(f"set_language failed for '{new_lang}'")

    def _rebuild_ui_for_language(self, selected_tab_key):
        """Destroys and recreates major UI parts to apply new language."""
        # This will now correctly log the NEW language (e.g., 'ar')
        logging.info(f"Rebuilding UI for language: {translations.CURRENT_LANGUAGE}. Selected key: {selected_tab_key}")

        # --- The rest of this method is correct and does not need changes ---
        if hasattr(self, 'header_frame') and self.header_frame.winfo_exists():
            self.header_frame.destroy()
        if hasattr(self, 'tabview') and self.tabview.winfo_exists():
            self.tabview.destroy()
            del self.tabview
            if hasattr(self, 'tab_upload'): del self.tab_upload
            if hasattr(self, 'tab_manage'): del self.tab_manage
            if hasattr(self, 'tab_settings'): del self.tab_settings
            if hasattr(self, 'tab_admin'): del self.tab_admin
        if hasattr(self, 'status_frame') and self.status_frame.winfo_exists():
            self.status_frame.destroy()

        self.header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent", height=60)
        self.header_frame.grid(row=0, column=0, sticky="ew", padx=15, pady=(10, 5))
        self.header_frame.grid_columnconfigure(0, weight=1)
        self.header_frame.grid_columnconfigure(1, weight=0)
        self.header_frame.grid_columnconfigure(2, weight=0)
        self.header_frame.grid_columnconfigure(3, weight=0)
        ctk.CTkLabel(self.header_frame, text=self._t("ctklabel_text_file_archiving_system"), font=("Segoe UI", 26, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.user_label = ctk.CTkLabel(self.header_frame, text="", font=("Segoe UI", 12))
        self.user_label.grid(row=0, column=1, padx=10, sticky="e")
        if self.current_user:
             user_text = f"{self._t('label_user')}: {self.current_user['username']} ({self._t('role_' + self.current_user['role'])})"
             self.user_label.configure(text=user_text)
        self.language_button = ctk.CTkButton(self.header_frame, text=self._t("button_text_switch_language"), command=self.switch_language, font=("Segoe UI", 12), width=120)
        self.language_button.grid(row=0, column=2, padx=10, sticky="e")
        self.logout_button_placeholder = ctk.CTkFrame(self.header_frame, fg_color="transparent", width=110)
        self.logout_button_placeholder.grid(row=0, column=3, padx=(5, 0), sticky="e")

        self.tabview = ctk.CTkTabview(self.main_frame, corner_radius=8)
        self.tabview.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))
        self._create_tabs()

        if self.current_user:
            logging.info("Setting up tab content after language switch...")
            try:
                if hasattr(self, 'tab_upload') and self.tab_upload.winfo_exists():
                     self.setup_upload_tab()
                if hasattr(self, 'tab_manage') and self.tab_manage.winfo_exists():
                     self.setup_manage_tab()
                if hasattr(self, 'tab_settings') and self.tab_settings.winfo_exists():
                     self.setup_settings_tab()
                if self.current_user["role"] == "admin":
                     admin_tab_name = self._t("tab_admin")
                     try:
                          self.tab_admin = self.tabview.nametowidget(self.tabview._w + '.inner.' + admin_tab_name)
                     except KeyError:
                          self.tab_admin = self.tabview.add(admin_tab_name)
                          self.tab_admin.grid_columnconfigure(0, weight=1)
                          self.tab_admin.grid_rowconfigure(0, weight=1)
                     if hasattr(self, 'tab_admin') and self.tab_admin.winfo_exists():
                          self.setup_admin_tab()
            except Exception as e:
                logging.error(f"Error re-setting up tabs after language switch: {e}", exc_info=True)
                messagebox.showerror(self._t("error_title"), f"{self._t('error_rebuild_tab_content')}: {e}")

        self.status_frame = ctk.CTkFrame(self.main_frame, height=30, corner_radius=0)
        self.status_frame.grid(row=2, column=0, sticky="ew", padx=0, pady=0)
        self.status_frame.grid_columnconfigure(0, weight=1)
        self.notification_label = ctk.CTkLabel(self.status_frame, text=self._t("ctklabel_text_ready"), font=("Segoe UI", 12))
        self.notification_label.grid(row=0, column=0, sticky="w", padx=10)
        self.progress_bar = ctk.CTkProgressBar(self.status_frame, width=150)
        self.progress_bar.grid(row=0, column=1, padx=10, sticky="e")
        self.progress_bar.set(0)

        self.add_logout_button()

        if selected_tab_key:
            try:
                new_label = self._t(selected_tab_key)
                logging.info(f"Attempting to re-select tab using new label: '{new_label}' (from key: '{selected_tab_key}')")
                if new_label in self.tabview._name_list:
                    self.main_app.after(50, lambda name=new_label: self.tabview.set(name))
                    logging.info(f"Scheduled selection of tab: '{new_label}'")
                else:
                    logging.warning(f"Tab with label '{new_label}' not found. Falling back to first tab.")
                    self._select_first_tab()
            except Exception as e:
                logging.error(f"Could not re-select tab for key '{selected_tab_key}'. Error: {e}", exc_info=True)
                self._select_first_tab()
        else:
            logging.warning("No selected_tab_key provided, selecting first available tab.")
            self._select_first_tab()
            
    # =========================================================================
    # ==== END OF LANGUAGE SWITCHING ==========================================
    # =========================================================================

    def load_translations(filename="translations.json"):
        global TRANSLATIONS, CURRENT_LANGUAGE
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path = os.path.join(script_dir, filename)
            if not os.path.exists(file_path):
                logging.error(f"Translation file not found at expected path: {file_path}")
                TRANSLATIONS = {"en": {}, "ar": {}}
                return TRANSLATIONS.get(CURRENT_LANGUAGE, {})
            
            with open(file_path, 'r', encoding='utf-8') as f:
                TRANSLATIONS = json.load(f)
                
            if "en" not in TRANSLATIONS:
                TRANSLATIONS["en"] = {}
            if "ar" not in TRANSLATIONS:
                TRANSLATIONS["ar"] = {}
                
            logging.info(f"Translations loaded successfully from {file_path}. Current language: {CURRENT_LANGUAGE}")
            return TRANSLATIONS.get(CURRENT_LANGUAGE, {})
            
        except FileNotFoundError:
            logging.error(f"Translation file '{filename}' not found.")
            TRANSLATIONS = {"en": {}, "ar": {}}
            return TRANSLATIONS.get(CURRENT_LANGUAGE, {})
        except json.JSONDecodeError:
            logging.error(f"Error decoding JSON from translation file '{filename}'.")
            TRANSLATIONS = {"en": {}, "ar": {}}
            return TRANSLATIONS.get(CURRENT_LANGUAGE, {})
        except Exception as e:
            logging.error(f"An unexpected error occurred loading translations: {e}", exc_info=True)
            TRANSLATIONS = {"en": {}, "ar": {}}
            return TRANSLATIONS.get(CURRENT_LANGUAGE, {})
    # Add this helper method inside FileArchiveApp class
    # def load_app_translations(self): # Removed
    #     """Loads translations for the current language into the instance.""" # Removed
    #     global CURRENT_LANGUAGE # Removed
    #     try: # Removed
    #         # Assuming load_translations returns the dict for the language # Removed
    #         self.translations = load_translations(CURRENT_LANGUAGE) # Removed
    #         if not self.translations: # Removed
    #             logging.warning(f"Loaded translations for '{CURRENT_LANGUAGE}' appear empty.") # Removed
    #         else: # Removed
    #             logging.info(f"Loaded {len(self.translations)} translations for '{CURRENT_LANGUAGE}'.") # Removed
    #     except Exception as e: # Removed
    #         logging.error(f"Failed to load translations for language '{CURRENT_LANGUAGE}': {e}", exc_info=True) # Removed
    #         self.translations = {} # Ensure it's an empty dict on failure # Removed

    # Add this helper method inside FileArchiveApp class
    def _create_tabs(self):
        """Creates the base tabs in the tabview using current translations."""
        # Clear existing tabs first (important for rebuild)
        if hasattr(self, 'tabview') and self.tabview.winfo_exists():
            # Store names before deleting to avoid modifying list while iterating
            current_tab_names = list(self.tabview._name_list) # Use internal list
            for tab_name in current_tab_names:
                try:
                    self.tabview.delete(tab_name)
                except Exception as e:
                    logging.warning(f"Could not delete tab '{tab_name}' during rebuild: {e}")
            # Clear references to old tab objects to prevent errors
            if hasattr(self, 'tab_upload'): del self.tab_upload
            if hasattr(self, 'tab_manage'): del self.tab_manage
            if hasattr(self, 'tab_settings'): del self.tab_settings
            if hasattr(self, 'tab_admin'): del self.tab_admin # If it existed

        # Add tabs using translated names from the current language via _t helper
        self.tab_upload = self.tabview.add(self._t("tab_upload_files"))
        self.tab_manage = self.tabview.add(self._t("tab_manage_files"))
        self.tab_settings = self.tabview.add(self._t("tab_settings"))

        # Configure tab grids (apply to all base tabs)
        for tab in [self.tab_upload, self.tab_manage, self.tab_settings]:
            if tab and tab.winfo_exists(): # Check if tab was created successfully
                tab.grid_columnconfigure(0, weight=1)
                tab.grid_rowconfigure(0, weight=1) # Allow content to expand vertically

        logging.debug(f"Base tabs created: {[self._t('tab_upload_files'), self._t('tab_manage_files'), self._t('tab_settings')]}")

    # Add this helper method inside FileArchiveApp class
    def _select_first_tab(self):
        """Helper to select the first available tab after UI rebuild."""
        try:
            # Get the list of currently displayed tab names
            tab_names = self.tabview._name_list
            if tab_names:
                first_tab_name = tab_names[0]
                logging.info(f"Selecting first available tab: '{first_tab_name}'")
                # Use after to ensure tabview is ready
                self.main_app.after(50, lambda name=first_tab_name: self.tabview.set(name))
            else:
                logging.warning("No tabs found to select after UI rebuild.")
        except Exception as e:
            logging.error(f"Error selecting first tab: {e}", exc_info=True)
    # --------------------------------------------------------------------------
    # Toggle Theme (Dark/Light)
    # --------------------------------------------------------------------------
    def toggle_theme(self, event=None):
        current_mode = ctk.get_appearance_mode()
        new_mode = "light" if current_mode == "dark" else "dark"
        ctk.set_appearance_mode(new_mode)
        self.ui_queue.put(lambda: messagebox.showinfo("Theme Changed", f"Switched to {new_mode.capitalize()} Mode"))
        logging.info(f"Theme toggled to {new_mode}")


    # --------------------------------------------------------------------------
    # Process UI Queue (One callback per cycle)
    # --------------------------------------------------------------------------
    def process_ui_queue(self):
        """Process UI Queue (One callback per cycle) with error handling"""
        try:
            callback = self.ui_queue.get_nowait()
            try:
                # Only run callback if main window exists
                if self.main_app.winfo_exists():
                    callback()
            except Exception as e:
                logging.error(f"Error in UI callback: {e}")
            finally:
                self.ui_queue.task_done()
        except queue.Empty:
            pass

        # Schedule next check only if main window still exists
        if hasattr(self, 'main_app') and self.main_app.winfo_exists():
            self.main_app.after(100, self.process_ui_queue)


    # --------------------------------------------------------------------------
    # Update Options for Header/Subheader (sets default subheader)
    # --------------------------------------------------------------------------
        # Place this method within the FileArchiveApp class
    def get_dynamic_folder_options(self, base_folder_path, template_options):
        """
        Gets folder options by merging template definition with actual folders on disk.

        Args:
            base_folder_path (str): The directory path to scan for folders.
            template_options (list | dict | None): Options from self.structure.
                                                   If dict, keys are used. If list, items are used.

        Returns:
            list: A sorted list of unique folder names found in template and on disk.
        """
        disk_folders = set()
        template_folders = set()

        # 1. Get folders from disk
        if os.path.isdir(base_folder_path):
            try:
                for item in os.listdir(base_folder_path):
                    # Check if it's a directory and not a hidden file/folder (optional)
                    if os.path.isdir(os.path.join(base_folder_path, item)) and not item.startswith('.'):
                        disk_folders.add(item)
                logging.debug(f"Disk scan found folders in '{base_folder_path}': {disk_folders}")
            except OSError as e:
                logging.warning(f"Could not scan directory '{base_folder_path}': {e}")
        else:
            logging.debug(f"Base folder path for dynamic options does not exist: {base_folder_path}")


        # 2. Get folders from template
        if isinstance(template_options, dict):
            template_folders = set(template_options.keys())
        elif isinstance(template_options, list):
            template_folders = set(template_options)
        # Ignore if template_options is None or other type

        logging.debug(f"Template options for '{base_folder_path}': {template_folders}")

        # 3. Merge and sort
        combined_options = sorted(list(template_folders.union(disk_folders)))
        logging.debug(f"Combined options for '{base_folder_path}': {combined_options}")

        return combined_options
    # --- Modify update_section_options_upload ---
    def update_section_options_upload(self, *args):
        if not hasattr(self, 'section_menu'): return # Safety check

        company_display_name = self.company_entry.get().strip()
        header_value = self.header_var.get()
        subheader_value = self.subheader_var.get() # Current subheader selection

        sections = [] # Default empty

        if company_display_name and header_value and subheader_value:
            try:
                safe_company_name = self.sanitize_path(company_display_name)
                header_template_options = self.structure.get(header_value, None)

                # Check if template structure is nested (dict)
                if isinstance(header_template_options, dict):
                    # Get template sections for this subheader
                    section_template_options = header_template_options.get(subheader_value, None)

                    # Determine disk path for sections
                    company_subheader_path = os.path.join(self.archives_path, safe_company_name, header_value, subheader_value)

                    # Get dynamic options (sections)
                    sections = self.archive_controller.get_dynamic_folder_options(company_subheader_path, section_template_options)
                    logging.info(f"Updating Section options for Company '{safe_company_name}', Path '{header_value}/{subheader_value}'. Found: {sections}")
                else:
                    # Flat structure - no sections defined in template or dynamically scanned at this level
                    logging.debug(f"update_section_options_upload: Header '{header_value}' is flat, no sections applicable.")
                    sections = [] # Explicitly clear

            except Exception as e:
                 logging.error(f"Error in update_section_options_upload for company '{company_display_name}': {e}", exc_info=True)
                 sections = [] # Clear on error

        self.section_menu.configure(values=sections)
        current_section = self.section_var.get()

        if sections:
            if current_section not in sections:
                self.section_var.set(sections[0])
            # else: keep current valid selection
        else:
            self.section_var.set("")

        # The trace on section_var will call update_subsection_options_upload

    # --- Modify update_subsection_options_upload ---
    def update_subsection_options_upload(self, *args):
        if not hasattr(self, 'subsection_menu'): return # Safety check

        company_display_name = self.company_entry.get().strip()
        header_value = self.header_var.get()
        subheader_value = self.subheader_var.get()
        section_value = self.section_var.get() # Current section selection

        subsections = [] # Default empty

        if company_display_name and header_value and subheader_value and section_value:
             try:
                safe_company_name = self.sanitize_path(company_display_name)
                header_template_options = self.structure.get(header_value, None)

                # Check if template structure is nested down to sections
                subsection_template_options = None
                if isinstance(header_template_options, dict):
                    section_dict = header_template_options.get(subheader_value, None)
                    if isinstance(section_dict, dict):
                        subsection_template_options = section_dict.get(section_value, None) # Should be a list or None

                # Determine disk path for subsections
                company_section_path = os.path.join(self.archives_path, safe_company_name, header_value, subheader_value, section_value)

                # Get dynamic options (subsections)
                # Pass the template list (or None)
                subsections = self.archive_controller.get_dynamic_folder_options(company_section_path, subsection_template_options)
                logging.info(f"Updating Subsection options for Company '{safe_company_name}', Path '{header_value}/{subheader_value}/{section_value}'. Found: {subsections}")

             except Exception as e:
                 logging.error(f"Error in update_subsection_options_upload for company '{company_display_name}': {e}", exc_info=True)
                 subsections = [] # Clear on error

        self.subsection_menu.configure(values=subsections)
        current_subsection = self.subsection_var.get()

        if subsections:
            if current_subsection not in subsections:
                 # Check if template defined a default and it exists, otherwise use first dynamic
                 default_in_template = subsection_template_options[0] if isinstance(subsection_template_options, list) and subsection_template_options else None
                 if default_in_template and default_in_template in subsections:
                      self.subsection_var.set(default_in_template)
                 else:
                      self.subsection_var.set(subsections[0])
            # else: keep current valid selection
        else:
            self.subsection_var.set("")

    # --- Modify update_options (Top Level) ---
    def update_options(self, *args):
        # --- Get Current Company ---
        company_display_name = self.company_entry.get().strip()
        if not company_display_name:
            # Clear dependent menus if no company selected
            self.subheader_menu.configure(values=[])
            self.subheader_var.set("")
            if hasattr(self, 'section_menu'):
                 self.section_menu.configure(values=[])
                 self.section_var.set("")
            if hasattr(self, 'subsection_menu'):
                 self.subsection_menu.configure(values=[])
                 self.subsection_var.set("")
            logging.debug("update_options: No company selected, clearing lower menus.")
            return

        try:
            # Get safe name - re-run create_structure to be safe if needed,
            # or assume it exists if we got this far. Let's assume it.
            # A better way might be to store current safe name when company is entered/changed.
            safe_company_name = self.sanitize_path(company_display_name) # Re-sanitize
            # Ensure self.current_company is updated if needed, maybe add a separate company change handler
            if not hasattr(self, 'current_company') or self.current_company.get("safe_name") != safe_company_name:
                 # This might re-run structure creation, which is ok.
                 self.create_company_structure(company_display_name)
                 safe_company_name = self.current_company["safe_name"]


            header_value = self.header_var.get()
            header_template_options = self.structure.get(header_value, None) # Get template structure for this header

            # --- Determine Path and Get Dynamic Options for Subheaders ---
            company_header_path = os.path.join(self.archives_path, safe_company_name, header_value)
            subh_options = self.archive_controller.get_dynamic_folder_options(company_header_path, header_template_options)

            logging.info(f"Updating Subheader options for Company '{safe_company_name}', Header '{header_value}'. Found: {subh_options}")

            self.subheader_menu.configure(values=subh_options)
            current_subheader = self.subheader_var.get()

            if subh_options:
                if current_subheader not in subh_options:
                    self.subheader_var.set(subh_options[0])
                # else: keep current valid selection
            else:
                self.subheader_var.set("")

            # The trace on subheader_var will call update_section_options_upload automatically
            # which will in turn call update_subsection_options_upload.

        except Exception as e:
             logging.error(f"Error in update_options for company '{company_display_name}': {e}", exc_info=True)
             # Optionally clear menus on error
             self.subheader_menu.configure(values=[])
             self.subheader_var.set("")

    # --------------------------------------------------------------------------
    # Login Dialog (Single mainloop via wait_window)
    # --------------------------------------------------------------------------
    def add_logout_button(self):
        """Add logout button to the header placeholder."""
        try:
            # Find the placeholder frame in the header
            if hasattr(self, 'logout_button_placeholder') and self.logout_button_placeholder.winfo_exists():
                # Clear any previous content in the placeholder (important for language switch)
                for widget in self.logout_button_placeholder.winfo_children():
                    widget.destroy()

                # Add the logout button directly into the placeholder frame
                logout_btn = ctk.CTkButton(self.logout_button_placeholder, text=self._t("ctkbutton_text_logout"),
                                        command=self.logout,
                                        font=("Segoe UI", 13), # Match header style
                                        height=35,             # Match header style
                                        width=100,             # Explicit width
                                        fg_color="#dc3545", hover_color="#c82333")
                # Pack inside the placeholder, aligned right if desired
                logout_btn.pack(side="right", padx=0, pady=0)
                logging.info("Logout button added to header placeholder.")
            else:
                logging.error("Could not find or access logout button placeholder in header.")
        except Exception as e:
            logging.error(f"Error adding logout button to header: {e}", exc_info=True)
    def authenticate_user(self):
        """Improved login dialog with better UI"""
        login_win = ctk.CTkToplevel(self.main_app)
        login_win.transient(self.main_app) # Stay on top of main window
        login_win.title("Login")
        login_win.resizable(False, False)
        self.center_window(login_win, 400, 300)
        login_win.grab_set() # Focus and block main window

        # Login title with larger font
        ctk.CTkLabel(login_win, text=get_translation("ctklabel_text_file_archiving_system"),
                    font=("Segoe UI", 24, "bold")).pack(pady=(20, 5))

        ctk.CTkLabel(login_win, text=get_translation("ctklabel_text_please_login"),
                    font=("Segoe UI", 16)).pack(pady=(0, 20))

        # Frame for inputs
        login_frame = ctk.CTkFrame(login_win)
        login_frame.pack(padx=20, pady=10, fill="x")

        # Username field
        ctk.CTkLabel(login_frame, text=get_translation("ctklabel_text_username"),
                    font=("Segoe UI", 14)).pack(anchor="w", padx=10, pady=(10, 0))

        username_entry = ctk.CTkEntry(login_frame, placeholder_text=get_translation("ctkentry_placeholder_text_enter_your_username"),
                                    width=300, font=("Segoe UI", 14))
        username_entry.pack(padx=10, pady=(5, 10), fill="x")

        # Password field
        ctk.CTkLabel(login_frame, text=get_translation("ctklabel_text_password"),
                    font=("Segoe UI", 14)).pack(anchor="w", padx=10, pady=(10, 0))

        password_entry = ctk.CTkEntry(login_frame, placeholder_text=get_translation("ctkentry_placeholder_text_enter_your_password"),
                                    show="*", width=300, font=("Segoe UI", 14))
        password_entry.pack(padx=10, pady=(5, 15), fill="x")

        # Remember the input field values (in a real app, you'd want secure handling)
        remember_var = ctk.BooleanVar(value=False)
        remember_check = ctk.CTkCheckBox(login_win, text=get_translation("ctkcheckbox_text_remember_username"),
                                        variable=remember_var, font=("Segoe UI", 12))
        remember_check.pack(pady=5)

        def do_login():
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            try:
                if username in users and pbkdf2_sha256.verify(password, users[username]["password"]):
                    self.current_user = {"username": username, "role": users[username]["role"]}
                    messagebox.showinfo("Login Success", f"Welcome {self.current_user['role'].capitalize()}!")
                    logging.info(f"{self.current_user['role'].capitalize()} '{username}' logged in.")

                    # Update the user label
                    if hasattr(self, 'user_label'):
                        self.user_label.configure(text=f"User: {username} ({self.current_user['role']})")

                    login_win.destroy()

                    # Setup the tab content BEFORE showing the main window
                    if hasattr(self, 'setup_upload_tab'):
                        self.setup_upload_tab()
                    if hasattr(self, 'setup_manage_tab'):
                        self.setup_manage_tab()
                    if hasattr(self, 'setup_settings_tab'):
                        self.setup_settings_tab()

                    # Show the admin tab if admin user
                    if self.current_user["role"] == "admin":
                        self.show_archive_folder()
                        if hasattr(self, 'tabview') and not hasattr(self, 'tab_admin'):
                            self.tab_admin = self.tabview.add(get_translation("tab_admin"))
                            self.setup_admin_tab()
                    else:
                        # Make sure the archive folder is hidden for non-admin users
                        self.hide_archive_folder()

                    # Add logout button
                    self.add_logout_button()

                    # Now show the main window after tabs are set up
                    self.main_app.deiconify()
                else:
                    messagebox.showerror("Login Error", "Invalid credentials!")
                    logging.warning("Failed login attempt.")
            except Exception as e:
                messagebox.showerror("Login Error", f"Error during login: {e}")
                logging.error(f"Login error: {e}")

        # Login button
        ctk.CTkButton(login_win, text=get_translation("ctkbutton_text_login"), command=do_login,
                    width=200, height=40, font=("Segoe UI", 16, "bold"),
                    fg_color="#2D7FF9", hover_color="#1A6CD6").pack(pady=15)

        # Focus on username entry
        username_entry.focus_set()

        # Bind Enter key to login button
        login_win.bind("<Return>", lambda event: do_login())

        login_win.grab_set()
        self.main_app.wait_window(login_win)
    def change_password(self):
        """Allow users to change their password"""
        if not self.current_user:
            return

        change_pwd_win = ctk.CTkToplevel(self.main_app)
        change_pwd_win.transient(self.main_app) # Stay on top
        change_pwd_win.title("Change Password")
        self.center_window(change_pwd_win, 400, 250)
        change_pwd_win.grab_set()

        ctk.CTkLabel(change_pwd_win, text=get_translation("ctklabel_text_change_password"),
                     font=("Segoe UI", 20, "bold")).pack(pady=15)

        # Current password
        current_pwd = ctk.CTkEntry(change_pwd_win, placeholder_text=get_translation("ctkentry_placeholder_text_current_password"),
                                 show="*", width=250, font=("Segoe UI", 14))
        current_pwd.pack(pady=10)

        # New password
        new_pwd = ctk.CTkEntry(change_pwd_win, placeholder_text=get_translation("ctkentry_placeholder_text_new_password"),
                              show="*", width=250, font=("Segoe UI", 14))
        new_pwd.pack(pady=10)

        # Confirm new password
        confirm_pwd = ctk.CTkEntry(change_pwd_win, placeholder_text=get_translation("ctkentry_placeholder_text_confirm_new_password"),
                                  show="*", width=250, font=("Segoe UI", 14))
        confirm_pwd.pack(pady=10)

        def change_password(self):
            """Only allow admin users to change their own password"""
            if not self.current_user:
                return

            # If the current user is not an admin, disallow password change
            if self.current_user["role"].lower() != "admin":
                messagebox.showerror("Permission Denied", "Only admin users can change their password.", parent=self.main_app)
                return

            change_pwd_win = ctk.CTkToplevel(self.main_app)
            change_pwd_win.transient(self.main_app)
            change_pwd_win.title("Change Password")
            self.center_window(change_pwd_win, 400, 250)
            change_pwd_win.grab_set()

            ctk.CTkLabel(change_pwd_win, text=get_translation("ctklabel_text_change_password"), font=("Segoe UI", 20, "bold")).pack(pady=15)

            # Current password
            current_pwd = ctk.CTkEntry(change_pwd_win, placeholder_text=get_translation("ctkentry_placeholder_text_current_password"), show="*", width=250, font=("Segoe UI", 14))
            current_pwd.pack(pady=10)

            # New password
            new_pwd = ctk.CTkEntry(change_pwd_win, placeholder_text=get_translation("ctkentry_placeholder_text_new_password"), show="*", width=250, font=("Segoe UI", 14))
            new_pwd.pack(pady=10)

            # Confirm new password
            confirm_pwd = ctk.CTkEntry(change_pwd_win, placeholder_text=get_translation("ctkentry_placeholder_text_confirm_new_password"), show="*", width=250, font=("Segoe UI", 14))
            confirm_pwd.pack(pady=10)

            def do_change_password():
                curr = current_pwd.get().strip()
                new = new_pwd.get().strip()
                confirm = confirm_pwd.get().strip()
                username = self.current_user['username']

                if not (curr and new and confirm):
                    messagebox.showerror("Error", "All fields are required", parent=change_pwd_win)
                    return

                if new != confirm:
                    messagebox.showerror("Error", "New passwords do not match", parent=change_pwd_win)
                    return

                if not pbkdf2_sha256.verify(curr, users[username]["password"]):
                    messagebox.showerror("Error", "Current password is incorrect", parent=change_pwd_win)
                    return

                users[username]["password"] = pbkdf2_sha256.hash(new)
                messagebox.showinfo("Success", "Password changed successfully", parent=change_pwd_win)
                logging.info(f"Admin '{username}' changed their password")
                change_pwd_win.destroy()

            # Change button
            ctk.CTkButton(change_pwd_win, text=get_translation("ctkbutton_text_change_password"),
                        command=do_change_password, width=200,
                        font=("Segoe UI", 14, "bold")).pack(pady=15)
    # --------------------------------------------------------------------------
    # Center Window Utility
    # --------------------------------------------------------------------------
    def center_window(self, window, width, height):
        window.update_idletasks()
        sw = window.winfo_screenwidth()
        sh = window.winfo_screenheight()
        x = (sw - width) // 2
        y = (sh - height) // 2
        window.geometry(f"{width}x{height}+{x}+{y}")

    # --------------------------------------------------------------------------
    # Archive Folder Management (Basic security by obscurity; not robust for production)
    # --------------------------------------------------------------------------
    def create_archive_folder(self):
        if not os.path.exists(self.archives_path):
            os.makedirs(self.archives_path)
            logging.info(f"Archives folder created: {self.archives_path}")
        self.secure_archive_folder()

    def show_archive_folder(self):
        if platform.system() == "Windows":
            subprocess.run(['attrib', '-h', '-s', '/S', '/D', self.archives_path], shell=True)
        else:
            subprocess.run(['chmod', '755', self.archives_path])
        logging.info(f"Archive folder shown: {self.archives_path}")

    def hide_archive_folder(self):
        """Hide the archive folder (non-recursively) with detailed logging"""
        start_time = time.time()  # Start timer
        try:
            archive_path = os.path.abspath(self.archives_path)
            if not os.path.exists(archive_path):
                logging.warning(f"Archive path does not exist: {archive_path}")
                return

            logging.info(f"Hiding archive folder (non-recursive) STARTED: {archive_path}")

            if platform.system() == "Windows":
                logging.info("Before Windows attrib command")
                subprocess.run(['attrib', '+h', '+s', archive_path], shell=True, check=True)
                logging.info("After Windows attrib command")
            else:
                logging.info("Before chmod command")
                subprocess.run(['chmod', '700', archive_path], check=True)
                logging.info("After chmod command")

            logging.info("Archive folder hidden successfully (non-recursive) ENDED")
        except Exception as e:
            logging.error(f"Error hiding archive folder: {e}")
        finally:
            end_time = time.time()  # End timer
            duration = end_time - start_time
            logging.info(f"hide_archive_folder() execution time: {duration:.4f} seconds") # Log duration

    def secure_archive_folder(self):
        # Note: This is minimal security and should be improved for sensitive data.
        if platform.system() == "Windows":
            subprocess.run(['attrib', '+h', '+s', '/S', '/D', self.archives_path], shell=True)
        else:
            subprocess.run(['chmod', '700', self.archives_path])
        logging.info(f"Archive folder secured: {self.archives_path}")

    # --------------------------------------------------------------------------
    # Admin Controls (Additional Functions for Admin Users)
    # --------------------------------------------------------------------------
    def add_admin_controls(self):
        if self.admin_controls_added:
            return
        refresh_btn = ctk.CTkButton(self.main_frame, text=get_translation("ctkbutton_text_refresh_folders"), command=self.refresh_folders, font=("Segoe UI", 14))
        refresh_btn.grid(row=9, column=0, pady=5)
        search_btn = ctk.CTkButton(self.main_frame, text=get_translation("ctkbutton_text_search_archive"), command=self.search_archive, font=("Segoe UI", 14))
        search_btn.grid(row=9, column=1, pady=5)
        dashboard_btn = ctk.CTkButton(self.main_frame, text=get_translation("ctkbutton_text_dashboard"), command=self.open_dashboard, font=("Segoe UI", 14))
        dashboard_btn.grid(row=10, column=0, columnspan=2, pady=5)
        self.admin_controls_added = True

    # --------------------------------------------------------------------------
    # Refresh Folders (Background Thread)
    # --------------------------------------------------------------------------
    def refresh_folders(self):
        def task():
            try:
                # Clear the ArchiveController cache first
                if hasattr(self, 'archive_controller'):
                    self.archive_controller.clear_cache()
                    logging.info("Admin triggered folder refresh, clearing full ArchiveController cache.")

                for root, dirs, _ in os.walk(self.archives_path):
                    for d in dirs:
                        show_path = os.path.join(root, d)
                        if platform.system() == "Windows":
                            subprocess.run(['attrib', '-h', '-s', '/S', '/D', show_path], shell=True)
                self.show_archive_folder()
                self.ui_queue.put(lambda: messagebox.showinfo("Refreshed", "All folders are now visible."))
                logging.info("Folders refreshed by admin.")
            except Exception as e:
                self.ui_queue.put(lambda: messagebox.showerror("Error", f"Failed to refresh folders: {e}"))
                logging.error(f"Error refreshing folders: {e}")
        threading.Thread(target=task, daemon=True).start()

    # --------------------------------------------------------------------------
    # Real-Time Monitoring Using Watchdog
    # --------------------------------------------------------------------------
    def start_monitoring(self):
        """Start real-time monitoring with proper error handling"""
        try:
            event_handler = ArchiveEventHandler(self.ui_queue, self.notification_label.configure)
            self.observer = Observer()
            self.observer.schedule(event_handler, self.archives_path, recursive=True)
            self.observer.start()
            logging.info("File monitoring started")
        except Exception as e:
            logging.error(f"Error starting file monitoring: {e}")
            messagebox.showwarning("Warning", "File monitoring could not be started.")

    # --------------------------------------------------------------------------
    # Helper: Reusable Selection Interface for Custom Dialogs
    # Returns (company_var, header_var, subheader_var, file_var)
    # --------------------------------------------------------------------------
    def create_selection_interface(self, parent):
        companies = [d for d in os.listdir(self.archives_path) if os.path.isdir(os.path.join(self.archives_path, d))]
        if not companies:
            messagebox.showinfo("Info", "No companies found in archive.")
            parent.destroy()
            return None, None, None, None, None # Return 5 Nones now

        # --- UI Elements ---
        ctk.CTkLabel(parent, text=get_translation("ctklabel_text_select_company"), font=("Segoe UI", 14)).pack(pady=5)
        company_var = ctk.StringVar(value=companies[0])
        company_menu = ctk.CTkOptionMenu(parent, variable=company_var, values=companies, font=("Segoe UI", 14))
        company_menu.pack(pady=5)

        headers = list(self.structure.keys())
        ctk.CTkLabel(parent, text=get_translation("ctklabel_text_select_header"), font=("Segoe UI", 14)).pack(pady=5)
        header_var = ctk.StringVar(value=headers[0])
        header_menu = ctk.CTkOptionMenu(parent, variable=header_var, values=headers, font=("Segoe UI", 14))
        header_menu.pack(pady=5)

        ctk.CTkLabel(parent, text=get_translation("ctklabel_text_select_subheader"), font=("Segoe UI", 14)).pack(pady=5)
        subheader_var = ctk.StringVar()
        subheader_menu = ctk.CTkOptionMenu(parent, variable=subheader_var, values=[], font=("Segoe UI", 14))
        subheader_menu.pack(pady=5)

        ctk.CTkLabel(parent, text=get_translation("ctklabel_text_select_section"), font=("Segoe UI", 14)).pack(pady=5)
        section_var = ctk.StringVar() # Use local var for this instance
        section_menu = ctk.CTkOptionMenu(parent, variable=section_var, values=[], font=("Segoe UI", 14))
        section_menu.pack(pady=5)

        # --- NEW SUBSECTION WIDGETS ---
        ctk.CTkLabel(parent, text=get_translation("ctklabel_text_select_subsection"), font=("Segoe UI", 14)).pack(pady=5)
        subsection_var = ctk.StringVar() # Use local var for this instance
        subsection_menu = ctk.CTkOptionMenu(parent, variable=subsection_var, values=[], font=("Segoe UI", 14))
        subsection_menu.pack(pady=5)
        # --- END NEW ---

        ctk.CTkLabel(parent, text=get_translation("ctklabel_text_select_file"), font=("Segoe UI", 14)).pack(pady=5)
        file_var = ctk.StringVar(value="")
        file_menu = ctk.CTkOptionMenu(parent, variable=file_var, values=[], font=("Segoe UI", 14))
        file_menu.pack(pady=5)

        # --- Update Logic ---
        # Define these *inside* create_selection_interface so they capture local vars

        def update_subsections_local(*args):
            opts = self.structure.get(header_var.get(), [])
            subsections = []
            if isinstance(opts, dict):
                section_dict = opts.get(subheader_var.get(), {})
                subsections = section_dict.get(section_var.get(), []) # Get subsections for selected section

            subsection_menu.configure(values=subsections)
            if subsections:
                if subsection_var.get() not in subsections:
                    subsection_var.set(subsections[0])
            else:
                subsection_var.set("")

        def update_sections_local(*args):
            opts = self.structure.get(header_var.get(), [])
            sections = []
            if isinstance(opts, dict):
                section_dict = opts.get(subheader_var.get(), {})
                sections = list(section_dict.keys()) # Sections are keys

            section_menu.configure(values=sections)
            if sections:
                if section_var.get() not in sections:
                    section_var.set(sections[0])
            else:
                section_var.set("")
            update_subsections_local() # Update subsections when section changes

        def update_subheaders_local(*args):
            opts = self.structure.get(header_var.get(), [])
            subh = []
            if isinstance(opts, dict):
                subh = list(opts.keys()) # Subheaders are keys
            else: # Flat structure
                subh = opts # Subheaders are the list items

            subheader_menu.configure(values=subh)
            if subh:
                if subheader_var.get() not in subh:
                    subheader_var.set(subh[0])
            else:
                subheader_var.set("")
            update_sections_local() # Update sections when subheader changes


        def update_file_menu_local(*args):
            folder = os.path.join(self.archives_path, company_var.get(), header_var.get())
            subh = subheader_var.get()
            sec = section_var.get()
            subsec = subsection_var.get() # Get subsection

            structure_options = self.structure.get(header_var.get(), [])

            if isinstance(structure_options, dict): # Nested Header
                if subh: folder = os.path.join(folder, subh)
                section_dict = structure_options.get(subh, {})
                if sec and sec in section_dict:
                     folder = os.path.join(folder, sec) # Add section path
                     # Check if this section HAS subsections before adding subsection path
                     if subsec and section_dict.get(sec, []): # Check if subsec selected AND section is defined to have them
                         folder = os.path.join(folder, subsec) # Add subsection path

            elif subh: # Flat structure (Header -> Subheader)
                folder = os.path.join(folder, subh)
                # No section/subsection for flat structure

            # --- Rest of file listing logic ---
            file_options = []
            if os.path.exists(folder):
                logging.debug(f"Listing files in: {folder}")
                try:
                    for f in os.listdir(folder):
                        if "_backup_" in f: # Skip backups in selection
                            continue
                        full_file_path = os.path.join(folder, f)
                        # Ensure it's actually a file before adding
                        if os.path.isfile(full_file_path):
                            file_options.append(f)
                        else:
                            logging.debug(f"Skipping non-file item: {f}")
                except Exception as e:
                    logging.error(f"Error listing files in {folder}: {e}")
            else:
                 logging.warning(f"File listing path does not exist: {folder}")

            file_menu.configure(values=file_options)
            file_var.set(file_options[0] if file_options else "")

        # --- Traces ---
        # Link traces to the *local* update functions defined above
        company_var.trace_add("write", update_file_menu_local)
        header_var.trace_add("write", lambda *args: (update_subheaders_local(), update_file_menu_local())) # Subheader updates section->subsection->file
        subheader_var.trace_add("write", lambda *args: (update_sections_local(), update_file_menu_local())) # Section updates subsection->file
        section_var.trace_add("write", lambda *args: (update_subsections_local(), update_file_menu_local())) # Subsection updates file
        subsection_var.trace_add("write", update_file_menu_local) # Subsection change directly updates file list

        # --- Initial Population ---
        update_subheaders_local() # This cascades down

        # Return the new subsection_var
        return company_var, header_var, subheader_var, section_var, subsection_var, file_var # Return 6 vars

    # --------------------------------------------------------------------------
    # Custom Preview Interface (No OS Explorer)
    # --------------------------------------------------------------------------
    def display_image(self, file_path):
        """Display an image with proper cleanup using CTkImage"""
        try:
            pil_image = Image.open(file_path)
            pil_image.thumbnail((350, 350)) # Keep PIL image for thumbnail
            ctk_image = ctk.CTkImage(pil_image, size=(350, 350)) # Create CTkImage
            return ctk_image # Return CTkImage object
        except Exception as e:
            logging.error(f"Error displaying image: {e}")
            return None

    
    def print_preview(self, file_path):
        """Handles printing functionality for the previewed image."""
        try:
            if platform.system() == "Windows":
                # Windows printing
                temp_dir = tempfile.gettempdir()
                temp_file_name = os.path.join(temp_dir, "archive_print.png")
                temp_file_name = os.path.abspath(temp_file_name)

                img = Image.open(file_path)
                img.save(temp_file_name, format='PNG')

                # Prompt the user for a printer name
                printer_name = choose_printer()
                if not printer_name:
                    messagebox.showerror("Print Error", "No printer selected.")
                    return

                # Construct mspaint command with the selected printer
                print_command = ['mspaint', '/pt', temp_file_name, printer_name]
                command_str = ' '.join(print_command)
                logging.info(f"Printing command: {command_str}")

                result = subprocess.run(print_command, capture_output=True, text=True, check=False)

                if result.returncode != 0:
                    error_message = f"Printing failed with code {result.returncode}. Stderr: {result.stderr}"
                    logging.error(error_message)
                    messagebox.showerror("Print Error", f"Printing failed. See log for details: {error_message}")
                else:
                    messagebox.showinfo("Print", "Printing started. Please check your printer.")

            elif platform.system() == "Darwin" or platform.system() == "Linux":
                # macOS and Linux printing (using lp command)
                subprocess.run(['lp', file_path], check=True)
                messagebox.showinfo("Print", "Printing started. Please check your printer.")
            else:
                messagebox.showinfo("Print", "Printing is not supported on this operating system.")
            logging.info(f"Print command executed for file: {file_path}")

        except Exception as e:
            messagebox.showerror("Print Error", f"Could not print file: {e}")
            logging.error(f"Error during print: {e}")
        finally:
            if platform.system() == "Windows" and 'temp_file_name' in locals():
                try:
                    os.remove(temp_file_name) # Clean up temp file after printing attempt
                except OSError as e:
                    logging.error(f"Error cleaning up temporary print file: {e}")


    def custom_preview_interface(self):
        """Opens a dialog to select and preview a file, including subsection path."""
        preview_win = ctk.CTkToplevel(self.main_app)
        preview_win.transient(self.main_app)
        preview_win.title("File Preview")
        self.center_window(preview_win, 450, 650) # Adjusted size
        preview_win.grab_set()

        # Create the selection interface within the preview window
        sel_vars = self.create_selection_interface(preview_win)
        if not sel_vars[0]: # Check if company selection was successful
             # create_selection_interface handles messages/closing if no companies
             return

        # Unpack all 6 variables
        company_var, header_var, subheader_var, section_var, subsection_var, file_var = sel_vars

        # Frame for preview button
        button_frame = ctk.CTkFrame(preview_win, fg_color="transparent")
        button_frame.pack(pady=20)

        def perform_preview():
            # Get all selected values
            comp = company_var.get()
            head = header_var.get()
            subh = subheader_var.get()
            sec = section_var.get()
            subsec = subsection_var.get() # Get selected subsection
            file_selected = file_var.get()

            if not comp or not head or not file_selected:
                messagebox.showerror("Error", "Please select Company, Header, and a File.", parent=preview_win)
                return

            # --- Build path including subsection conditionally ---
            folder = os.path.join(self.archives_path, comp, head)
            structure_options = self.structure.get(head, [])
            has_subsections_defined = False
            if isinstance(structure_options, dict): # Nested Header
                if subh: folder = os.path.join(folder, subh)
                section_dict = structure_options.get(subh, {})
                if sec and sec in section_dict:
                    folder = os.path.join(folder, sec)
                    subsections_list = section_dict.get(sec, [])
                    has_subsections_defined = bool(subsections_list)
                    # Add subsection path ONLY if subsection is selected AND section supports it
                    if subsec and has_subsections_defined:
                         folder = os.path.join(folder, subsec)
            elif subh: # Flat structure
                 folder = os.path.join(folder, subh)
            # --- End Path Building ---

            file_path = os.path.join(folder, file_selected)
            logging.info(f"[Preview] Attempting to preview: {file_path}")

            if not os.path.exists(file_path):
                 messagebox.showerror("Error", f"File not found:\n{file_path}", parent=preview_win)
                 logging.error(f"[Preview] File not found at path: {file_path}")
                 return

            ext = os.path.splitext(file_path)[1].lower()

            if ext in IMAGE_EXTENSIONS:
                # Display image in a new Toplevel window
                img_win = ctk.CTkToplevel(preview_win) # Toplevel of preview_win
                img_win.transient(preview_win)
                img_win.title(f"Preview: {file_selected}")
                img_win.grab_set() # Focus on image window
                img_win.geometry("600x550") # Adjust size as needed

                try:
                    # Use CTkImage for better integration
                    ctk_img = ctk.CTkImage(Image.open(file_path), size=(550, 450)) # Adjust display size
                    label = ctk.CTkLabel(img_win, image=ctk_img, text=get_translation("ctklabel_text_empty_string")) # Use text=get_translation("ctklabel_text_empty_string")
                    label.pack(pady=10, padx=10)

                    # Add Print button to image window
                    print_btn = ctk.CTkButton(img_win, text=get_translation("ctkbutton_text_print"),
                                              command=lambda p=file_path: self.print_preview(p),
                                              font=("Segoe UI", 14))
                    print_btn.pack(pady=(5, 10))
                    img_win.after(100, lambda: img_win.lift()) # Ensure it comes to front

                except Exception as e:
                    img_win.destroy() # Close the window on error
                    messagebox.showerror("Preview Error", f"Could not display image:\n{e}", parent=preview_win)
                    logging.error(f"[Preview] Error displaying image {file_path}: {e}", exc_info=True)
            else:
                # For non-image files, offer to open with the default OS application.
                if messagebox.askyesno("Open File", f"'{file_selected}' is not an image.\nDo you want to open it with the default application?", parent=preview_win):
                    try:
                        self.open_path(file_path)
                    except Exception as e:
                        messagebox.showerror("Open Error", f"Could not open file:\n{e}", parent=preview_win)
                        logging.error(f"[Preview] Error opening file {file_path} with OS: {e}", exc_info=True)

        preview_btn = ctk.CTkButton(button_frame, text=get_translation("ctkbutton_text_previewopen_selected_file"), command=perform_preview, font=("Segoe UI", 14))
        preview_btn.pack()

    def custom_rollback_interface(self):
        """Opens a dialog to select a file and rollback to a previous version, including subsection path."""
        rb_win = ctk.CTkToplevel(self.main_app)
        rb_win.transient(self.main_app)
        rb_win.title("Rollback File")
        self.center_window(rb_win, 450, 700) # Adjusted size for backup list
        rb_win.grab_set()

        # Create the selection interface
        sel_vars = self.create_selection_interface(rb_win)
        if not sel_vars[0]: return # Exit if no companies

        # Unpack all 6 variables
        company_var, header_var, subheader_var, section_var, subsection_var, file_var = sel_vars

        # Backup selection widgets
        ctk.CTkLabel(rb_win, text=get_translation("ctklabel_text_select_backup_version"), font=("Segoe UI", 14)).pack(pady=(10,2), anchor="w", padx=20)
        backup_var = ctk.StringVar(value="")
        backup_menu = ctk.CTkOptionMenu(rb_win, variable=backup_var, values=[""], font=("Segoe UI", 12), width=300)
        backup_menu.pack(pady=(0,10), padx=20, fill="x")

        # Frame for buttons
        button_frame = ctk.CTkFrame(rb_win, fg_color="transparent")
        button_frame.pack(pady=20)


        def update_backup_menu(*args):
            comp = company_var.get()
            head = header_var.get()
            subh = subheader_var.get()
            sec = section_var.get()
            subsec = subsection_var.get() # Get subsection
            selected_file = file_var.get()

            backup_options = []
            if comp and head and selected_file:
                # --- Build path including subsection ---
                folder = os.path.join(self.archives_path, comp, head)
                structure_options = self.structure.get(head, [])
                has_subsections_defined = False
                if isinstance(structure_options, dict):
                    if subh: folder = os.path.join(folder, subh)
                    section_dict = structure_options.get(subh, {})
                    if sec and sec in section_dict:
                        folder = os.path.join(folder, sec)
                        subsections_list = section_dict.get(sec, [])
                        has_subsections_defined = bool(subsections_list)
                        if subsec and has_subsections_defined:
                             folder = os.path.join(folder, subsec)
                elif subh:
                     folder = os.path.join(folder, subh)
                # --- End Path Building ---

                if os.path.exists(folder):
                    try:
                        base_name, ext = os.path.splitext(selected_file)
                        # Regex to find backups: base_backup_TIMESTAMP_optionalCount.ext
                        pattern = re.compile(re.escape(base_name) + r"_backup_(\d{14}(?:_\d+)?)" + re.escape(ext) + r"$")
                        all_files = os.listdir(folder)
                        for fname in all_files:
                            match = pattern.match(fname)
                            if match and os.path.isfile(os.path.join(folder, fname)):
                                backup_options.append(fname)

                        # Sort backups by timestamp (descending - newest first)
                        def sort_key(filename):
                            match = pattern.match(filename)
                            if match:
                                try:
                                    # Extract full timestamp part (e.g., 20230101120000 or 20230101120000_1)
                                    ts_part = match.group(1)
                                    # Use only the main 14 digits for primary sorting
                                    return int(ts_part[:14])
                                except (ValueError, IndexError):
                                    return 0 # Error case, sort to beginning
                            return 0 # Should not happen if pattern matched

                        backup_options.sort(key=sort_key, reverse=True)
                        logging.debug(f"[Rollback] Found backups for {selected_file} in {folder}: {backup_options}")

                    except Exception as e:
                        logging.error(f"[Rollback] Error finding/sorting backups for {selected_file}: {e}")
                else:
                    logging.warning(f"[Rollback] Folder not found for backup search: {folder}")

            backup_menu.configure(values=backup_options if backup_options else [""])
            if backup_options:
                 # Keep current selection if valid, else set to first
                 if backup_var.get() not in backup_options:
                      backup_var.set(backup_options[0])
            else:
                 backup_var.set("")

        # Trigger backup update when the main file selection changes
        file_var.trace_add("write", update_backup_menu)
        # Also trigger when path components change (as file list depends on them)
        company_var.trace_add("write", update_backup_menu)
        header_var.trace_add("write", update_backup_menu)
        subheader_var.trace_add("write", update_backup_menu)
        section_var.trace_add("write", update_backup_menu)
        subsection_var.trace_add("write", update_backup_menu)
        # Initial population
        update_backup_menu()


        def perform_rollback():
            comp = company_var.get()
            head = header_var.get()
            subh = subheader_var.get()
            sec = section_var.get()
            subsec = subsection_var.get()
            original_file = file_var.get()
            backup_file = backup_var.get()

            if not comp or not head or not original_file:
                 messagebox.showerror("Error", "Please select Company, Header, and File.", parent=rb_win)
                 return

            if not backup_file:
                 messagebox.showerror("Error", "No backup version selected.", parent=rb_win)
                 return

            # --- Build path including subsection ---
            folder = os.path.join(self.archives_path, comp, head)
            structure_options = self.structure.get(head, [])
            has_subsections_defined = False
            if isinstance(structure_options, dict):
                if subh: folder = os.path.join(folder, subh)
                section_dict = structure_options.get(subh, {})
                if sec and sec in section_dict:
                    folder = os.path.join(folder, sec)
                    subsections_list = section_dict.get(sec, [])
                    has_subsections_defined = bool(subsections_list)
                    if subsec and has_subsections_defined:
                         folder = os.path.join(folder, subsec)
            elif subh:
                 folder = os.path.join(folder, subh)
            # --- End Path Building ---

            original_path = os.path.join(folder, original_file)
            backup_path = os.path.join(folder, backup_file)

            if not os.path.exists(original_path):
                 messagebox.showerror("Error", f"Original file not found:\n{original_path}", parent=rb_win)
                 return
            if not os.path.exists(backup_path):
                 messagebox.showerror("Error", f"Selected backup file not found:\n{backup_path}", parent=rb_win)
                 return

            # Confirmation
            if not messagebox.askyesno("Confirm Rollback",
                                       f"Restore '{original_file}'\nfrom backup '{backup_file}'?\n\n"
                                       f"The current version of '{original_file}' will be overwritten.",
                                       icon='warning', parent=rb_win):
                return

            try:
                # Perform rollback by copying backup over original
                shutil.copy2(backup_path, original_path) # Use copy2 to preserve metadata
                messagebox.showinfo("Success", f"Rolled back '{original_file}'\nto version from '{backup_file}'", parent=rb_win)
                logging.info(f"[Rollback] Success: {original_path} restored from {backup_path}")
                rb_win.destroy()
            except Exception as e:
                messagebox.showerror("Rollback Error", f"Rollback failed: {e}", parent=rb_win)
                logging.error(f"[Rollback] Failed to copy {backup_path} to {original_path}: {e}")

        # Note: Removed the 'Delete' functionality from this button for clarity.
        # If delete is needed, it should be a separate function/button.
        rollback_btn = ctk.CTkButton(button_frame, text=get_translation("ctkbutton_text_rollback_to_selected_backup"), command=perform_rollback, font=("Segoe UI", 14))
        rollback_btn.pack()


    # --------------------------------------------------------------------------
    # Custom Rollback Interface (with Backup Sorting & Delete Option)
    # --------------------------------------------------------------------------
    def create_company_structure(self, company_name):
        """Create folder structure for a company with Unicode support, including subsections."""
        try:
            safe_company_name = self.sanitize_path(company_name)
            # Update current company info
            self.current_company = {"display_name": company_name, "safe_name": safe_company_name}

            base_path = os.path.join(self.archives_path, safe_company_name)
            logging.info(f"Creating/Verifying structure for company: {company_name} (Safe Path: {safe_company_name})")

            # Create base company folder
            os.makedirs(base_path, exist_ok=True)

            # Iterate through defined structure
            for header, sub_level_data in self.structure.items():
                header_path = os.path.join(base_path, header)
                os.makedirs(header_path, exist_ok=True)

                if isinstance(sub_level_data, dict): # Nested: Header -> Subheader -> Section -> Subsection
                    for subheader, section_dict in sub_level_data.items():
                        subheader_path = os.path.join(header_path, subheader)
                        os.makedirs(subheader_path, exist_ok=True)

                        if isinstance(section_dict, dict): # Ensure it's the expected Section dict
                            for section, subsection_list in section_dict.items():
                                section_path = os.path.join(subheader_path, section)
                                os.makedirs(section_path, exist_ok=True)

                                # Create subsection folders if defined
                                if isinstance(subsection_list, list):
                                    for subsection in subsection_list:
                                        if subsection: # Skip if subsection name is empty
                                             subsection_path = os.path.join(section_path, subsection)
                                             os.makedirs(subsection_path, exist_ok=True)
                                else:
                                     logging.warning(f"Expected list of subsections for {header}/{subheader}/{section}, found: {type(subsection_list)}")
                        else:
                             logging.warning(f"Expected dict for sections under {header}/{subheader}, found: {type(section_dict)}")

                elif isinstance(sub_level_data, list): # Flat: Header -> Subheader list
                    for subheader_item in sub_level_data:
                        if subheader_item: # Skip if item name is empty
                             subheader_path = os.path.join(header_path, subheader_item)
                             os.makedirs(subheader_path, exist_ok=True)
                else:
                     logging.warning(f"Unknown structure type for header '{header}': {type(sub_level_data)}")

            logging.debug(f"Structure verification complete for: {safe_company_name}")

        except Exception as e:
            logging.error(f"Failed to create company structure for '{company_name}': {e}", exc_info=True)
            # Re-raise the exception so the calling function knows structure creation failed
            raise

    # --------------------------------------------------------------------------
    # Open Path using OS Commands (for admin use if needed)
    # --------------------------------------------------------------------------
    def open_path(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open {path}: {e}")


    def setup_admin_tab(self):
        """Configure the admin tab with administrative functions"""
        # Create a tabview within the admin tab for better organization
        """Configure the admin tab with administrative functions"""
        # Create a tabview within the admin tab for better organization
        admin_tabview = ctk.CTkTabview(self.tab_admin)
        admin_tabview.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Get the names ONCE ---
        system_tab_name = get_translation("admin_system_tab")
        users_tab_name = get_translation("admin_users_tab")
        stats_tab_name = get_translation("admin_stats_tab")
        logging.debug(f"Admin Tab Names: System='{system_tab_name}', Users='{users_tab_name}', Stats='{stats_tab_name}'")

        # --- Add tabs using the retrieved names ---
        system_tab = admin_tabview.add(system_tab_name)
        users_tab = admin_tabview.add(users_tab_name)
        stats_tab = admin_tabview.add(stats_tab_name)
        logging.debug(f"Added admin sub-tabs.")

        # =====================================================================
        # System Management Tab
        # =====================================================================
        sys_frame = ctk.CTkFrame(system_tab)
        sys_frame.pack(fill="both", expand=True, padx=20, pady=15)

        # Header with icon
        header_frame = ctk.CTkFrame(sys_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(10, 20))

        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 24)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_system_management"),
                    font=("Segoe UI", 20, "bold")).pack(side="left")

        # Admin buttons grid
        admin_buttons = ctk.CTkFrame(sys_frame, fg_color="transparent")
        admin_buttons.pack(fill="x", pady=10)
        admin_buttons.grid_columnconfigure((0, 1), weight=1, uniform="column")

        # Refresh folders button with icon
        refresh_btn = ctk.CTkButton(admin_buttons, text=get_translation("ctkbutton_text_refresh_folders"),
                                    command=self.refresh_folders,
                                    font=("Segoe UI", 14),
                                    height=45,
                                    corner_radius=8)
        refresh_btn.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        # Search archive button with icon
        search_btn = ctk.CTkButton(admin_buttons, text=get_translation("ctkbutton_text_search_archive"),
                                command=self.search_archive,
                                font=("Segoe UI", 14),
                                height=45,
                                corner_radius=8)
        search_btn.grid(row=0, column=1, padx=10, pady=10, sticky="ew")

        # Dashboard button (full width) with icon
        dashboard_btn = ctk.CTkButton(sys_frame, text=get_translation("ctkbutton_text_open_admin_dashboard"),
                                    command=self.open_dashboard,
                                    font=("Segoe UI", 14, "bold"),
                                    height=45,
                                    corner_radius=8,
                                    fg_color="#2D7FF9", hover_color="#1A6CD6")
        dashboard_btn.pack(fill="x", pady=15)

        # System info section
        info_frame = ctk.CTkFrame(sys_frame)
        info_frame.pack(fill="x", pady=15)

        ctk.CTkLabel(info_frame, text=get_translation("ctklabel_text_system_information"),
                    font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=15, pady=(10, 5))

        # Get system info
        import platform
        system_info = f"OS: {platform.system()} {platform.version()}\n"
        system_info += f"Python: {platform.python_version()}\n"
        system_info += f"Archives Path: {self.archives_path}"

        info_text = ctk.CTkTextbox(info_frame, height=80, font=("Segoe UI", 13))
        info_text.pack(fill="x", padx=15, pady=(5, 10))
        info_text.insert("1.0", system_info)
        info_text.configure(state="disabled")

        # =====================================================================
        # User Management Tab
        # =====================================================================
        # Create a modern, card-based layout for user management
        user_mgmt_frame = ctk.CTkFrame(users_tab)
        user_mgmt_frame.pack(fill="both", expand=True, padx=20, pady=15)

        # Header with icon and add user button
        header_frame = ctk.CTkFrame(user_mgmt_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(10, 20))
        header_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 24)).grid(row=0, column=0, padx=(0, 10))
        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_user_management"),
                    font=("Segoe UI", 20, "bold")).grid(row=0, column=1, sticky="w")

        # Add user button with icon
        add_user_btn = ctk.CTkButton(header_frame, text=get_translation("ctkbutton_text_add_user"),
                                    command=self.add_user_dialog,
                                    font=("Segoe UI", 14),
                                    height=40,
                                    width=120,
                                    corner_radius=8,
                                    fg_color="#28a745", hover_color="#218838")
        add_user_btn.grid(row=0, column=2, padx=10)

        # Search and filter section
        filter_frame = ctk.CTkFrame(user_mgmt_frame, fg_color="transparent")
        filter_frame.pack(fill="x", pady=(0, 10))

        ctk.CTkLabel(filter_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 5))

        self.user_search_var = ctk.StringVar()
        user_search = ctk.CTkEntry(filter_frame, placeholder_text=get_translation("ctkentry_placeholder_text_search_users"),
                                width=200, font=("Segoe UI", 13),
                                textvariable=self.user_search_var)
        user_search.pack(side="left", padx=5)

        # Role filter
        self.role_filter_var = ctk.StringVar(value="All")
        ctk.CTkLabel(filter_frame, text=get_translation("ctklabel_text_role"), font=("Segoe UI", 13)).pack(side="left", padx=(15, 5))
        role_filter = ctk.CTkOptionMenu(filter_frame, values=["All", "Admin", "User"],
                                    variable=self.role_filter_var,
                                    width=100, font=("Segoe UI", 13))
        role_filter.pack(side="left", padx=5)

        # Connect search and filter to refresh function
        self.user_search_var.trace("w", lambda *args: self.refresh_user_list())
        self.role_filter_var.trace("w", lambda *args: self.refresh_user_list())

        # User list in a scrollable frame with modern styling
        user_list_frame = ctk.CTkScrollableFrame(user_mgmt_frame, height=350)
        user_list_frame.pack(fill="both", expand=True, pady=10)

        # Store reference to the frame for later updates
        self.user_list_frame = user_list_frame

        # Status bar at the bottom
        self.user_status_label = ctk.CTkLabel(user_mgmt_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 12))
        self.user_status_label.pack(anchor="w", padx=10, pady=(5, 10))

        # Populate the user list
        self.refresh_user_list()

        # =====================================================================
        # Statistics Tab
        # =====================================================================
        stats_frame = ctk.CTkFrame(stats_tab)
        stats_frame.pack(fill="both", expand=True, padx=20, pady=15)

        # Header with icon
        header_frame = ctk.CTkFrame(stats_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(10, 20))

        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 24)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_system_statistics"),
                    font=("Segoe UI", 20, "bold")).pack(side="left")

        # Simple stats display
        stats_info = ctk.CTkFrame(stats_frame)
        stats_info.pack(fill="x", pady=10)

        # Count users by role
        admin_count = sum(1 for user in users.values() if user["role"] == "admin")
        user_count = sum(1 for user in users.values() if user["role"] == "user")

        # User stats
        user_stats_frame = ctk.CTkFrame(stats_info)
        user_stats_frame.pack(fill="x", padx=15, pady=10)

        ctk.CTkLabel(user_stats_frame, text=get_translation("ctklabel_text_user_statistics"),
                    font=("Segoe UI", 16, "bold")).pack(anchor="w", pady=(5, 10))

        ctk.CTkLabel(user_stats_frame, text=f"Total Users: {len(users)}",
                    font=("Segoe UI", 14)).pack(anchor="w", padx=20, pady=2)
        ctk.CTkLabel(user_stats_frame, text=f"Admin Users: {admin_count}",
                    font=("Segoe UI", 14)).pack(anchor="w", padx=20, pady=2)
        ctk.CTkLabel(user_stats_frame, text=f"Regular Users: {user_count}",
                    font=("Segoe UI", 14)).pack(anchor="w", padx=20, pady=2)

    def refresh_user_list(self):
        """Refresh the user list display with search and filter"""
        # Return if user_list_frame doesn't exist anymore
        if not hasattr(self, 'user_list_frame') or not self.user_list_frame.winfo_exists():
            return

        # Clear existing widgets
        for widget in self.user_list_frame.winfo_children():
            widget.destroy()

        # Get search and filter values
        search_text = self.user_search_var.get().lower() if hasattr(self, 'user_search_var') else ""
        role_filter = self.role_filter_var.get() if hasattr(self, 'role_filter_var') else "All"

        # Filter users
        filtered_users = {}
        for username, details in users.items():
            # Apply search filter
            if search_text and search_text not in username.lower():
                continue

            # Apply role filter
            if role_filter != "All" and details["role"] != role_filter.lower():
                continue

            filtered_users[username] = details

        # Update status label
        if hasattr(self, 'user_status_label'):
            self.user_status_label.configure(text=f"Showing {len(filtered_users)} of {len(users)} users")

        # If no users found
        if not filtered_users:
            no_results = ctk.CTkLabel(self.user_list_frame,
                                    text="No users found matching your criteria",
                                    font=("Segoe UI", 14))
            no_results.pack(pady=50)
            return

        # Add user cards
        for username, details in filtered_users.items():
            # Create a card-like frame for each user
            card = ctk.CTkFrame(self.user_list_frame)
            card.pack(fill="x", padx=5, pady=5)

            # Main content frame
            content = ctk.CTkFrame(card, fg_color="transparent")
            content.pack(fill="x", padx=10, pady=10)
            content.grid_columnconfigure(1, weight=1)

            # User icon based on role
            icon = "🔑" if details["role"] == "admin" else "👤"
            ctk.CTkLabel(content, text=icon, font=("Segoe UI", 24)).grid(row=0, column=0, rowspan=2, padx=(0, 10))

            # Username with larger font
            ctk.CTkLabel(content, text=username,
                        font=("Segoe UI", 16, "bold")).grid(row=0, column=1, sticky="w")

            # Role with badge-like appearance
            role_frame = ctk.CTkFrame(content, fg_color="#2D7FF9" if details["role"] == "admin" else "#6c757d",
                                    corner_radius=5, height=22)
            role_frame.grid(row=1, column=1, sticky="w")
            role_frame.grid_propagate(False)

            ctk.CTkLabel(role_frame, text=details["role"].upper(),
                        font=("Segoe UI", 11, "bold"),
                        text_color="white").pack(padx=8, pady=0)

            # Action buttons
            action_frame = ctk.CTkFrame(content, fg_color="transparent")
            action_frame.grid(row=0, column=2, rowspan=2, padx=10, sticky="e")

            # Edit button
            edit_btn = ctk.CTkButton(action_frame, text=get_translation("ctkbutton_text_edit"),
                                    command=lambda u=username: self.edit_user_dialog(u),
                                    font=("Segoe UI", 12),
                                    width=80, height=30,
                                    corner_radius=6,
                                    fg_color="#ffc107", hover_color="#e0a800",
                                    text_color="#000000")
            edit_btn.pack(side="left", padx=5)

            # Delete button (don't allow deleting the current user)
            delete_btn = ctk.CTkButton(action_frame, text=get_translation("ctkbutton_text_delete"),
                                    command=lambda u=username: self.delete_user_confirm(u),
                                    font=("Segoe UI", 12),
                                    width=80, height=30,
                                    corner_radius=6,
                                    fg_color="#dc3545", hover_color="#c82333")
            delete_btn.pack(side="left", padx=5)

            # Disable delete button if it's the current user
            if self.current_user and username == self.current_user["username"]:
                delete_btn.configure(state="disabled")

    def add_user_dialog(self):
        """Show enhanced dialog to add a new user"""
        add_win = ctk.CTkToplevel(self.main_app)
        add_win.transient(self.main_app) # Stay on top
        add_win.title("Add New User")
        add_win.attributes("-topmost", True)
        self.center_window(add_win, 650, 600)
        add_win.grab_set()

        # Main frame with padding
        main_frame = ctk.CTkFrame(add_win)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header with icon
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 28)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_add_new_user"),
                    font=("Segoe UI", 22, "bold")).pack(side="left")

            # Form fields
        form_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        form_frame.pack(fill="x", pady=10)

        # Username field with icon
        username_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        username_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(username_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(username_frame, text=get_translation("ctklabel_text_username"),
                    font=("Segoe UI", 14, "bold")).pack(side="left")

        username_entry = ctk.CTkEntry(form_frame, placeholder_text=get_translation("ctkentry_placeholder_text_enter_username"),
                                    width=400, height=35, font=("Segoe UI", 14))
        username_entry.pack(pady=(0, 15))

        # Password field with icon
        password_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        password_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(password_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(password_frame, text=get_translation("ctklabel_text_password"),
                    font=("Segoe UI", 14, "bold")).pack(side="left")

        password_entry = ctk.CTkEntry(form_frame, placeholder_text=get_translation("ctkentry_placeholder_text_enter_password"),
                                    show="•", width=400, height=35, font=("Segoe UI", 14))
        password_entry.pack(pady=(0, 15))

        # Show/hide password toggle
        password_visible = False

        def toggle_password_visibility():
            nonlocal password_visible
            password_visible = not password_visible
            password_entry.configure(show="" if password_visible else "•")
            toggle_btn.configure(text="👁️ Hide" if password_visible else "👁️ Show")

        toggle_btn = ctk.CTkButton(form_frame, text=get_translation("ctkbutton_text_show"),
                                command=toggle_password_visibility,
                                font=("Segoe UI", 12),
                                width=80, height=25,
                                fg_color="#6c757d", hover_color="#5a6268")
        toggle_btn.pack(anchor="e", padx=5)

        # Role selection with styled radio buttons
        role_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        role_frame.pack(fill="x", pady=15)

        ctk.CTkLabel(role_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(role_frame, text=get_translation("ctklabel_text_user_role"),
                    font=("Segoe UI", 14, "bold")).pack(side="left")

        role_var = ctk.StringVar(value="user")

        role_options = ctk.CTkFrame(form_frame, fg_color="transparent")
        role_options.pack(fill="x", pady=(0, 15))

        # Admin radio with custom styling
        admin_frame = ctk.CTkFrame(role_options, fg_color="#f8f9fa", corner_radius=6)
        admin_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))

        admin_radio = ctk.CTkRadioButton(admin_frame, text=get_translation("ctkradiobutton_text_admin"),
                                        variable=role_var, value="admin",
                                        font=("Segoe UI", 14),
                                        border_width_checked=6,
                                        fg_color="#2D7FF9",
                                        hover_color="#1A6CD6")
        admin_radio.pack(side="left", padx=15, pady=10)

        ctk.CTkLabel(admin_frame, text=get_translation("ctklabel_text_full_system_access"),
                    font=("Segoe UI", 12),
                    text_color="#6c757d").pack(side="left", padx=5)

        # User radio with custom styling
        user_frame = ctk.CTkFrame(role_options, fg_color="#f8f9fa", corner_radius=6)
        user_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))

        user_radio = ctk.CTkRadioButton(user_frame, text=get_translation("ctkradiobutton_text_user"),
                                        variable=role_var, value="user",
                                        font=("Segoe UI", 14),
                                        border_width_checked=6,
                                        fg_color="#28a745",
                                        hover_color="#218838")
        user_radio.pack(side="left", padx=15, pady=10)

        ctk.CTkLabel(user_frame, text=get_translation("ctklabel_text_limited_access"),
                    font=("Segoe UI", 12),
                    text_color="#6c757d").pack(side="left", padx=5)

        # Buttons frame
        buttons_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", pady=(20, 0))

        # Cancel button
        cancel_btn = ctk.CTkButton(buttons_frame, text=get_translation("ctkbutton_text_cancel"),
                                command=add_win.destroy,
                                font=("Segoe UI", 14),
                                width=120, height=40,
                                fg_color="#6c757d", hover_color="#5a6268")
        cancel_btn.pack(side="left", padx=(0, 10))

        def do_add_user():
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            role = role_var.get()

            # Validate inputs
            if not username or not password:
                messagebox.showerror("Error", "Username and password are required", parent=add_win)
                return

            if username in users:
                messagebox.showerror("Error", f"User '{username}' already exists", parent=add_win)
                return

            # Add the new user
            users[username] = {
                "password": pbkdf2_sha256.hash(password),
                "role": role
            }

            # Save changes to file
            self.save_user_to_db(username, pbkdf2_sha256.hash(password), role)

            messagebox.showinfo("Success", f"User '{username}' added successfully", parent=add_win)
            logging.info(f"Admin '{self.current_user['username']}' added new user '{username}' with role '{role}'")

            # Refresh the user list
            self.refresh_user_list()
            add_win.destroy()

        # Add button
        add_btn = ctk.CTkButton(buttons_frame, text=get_translation("ctkbutton_text_add_user"),
                            command=do_add_user,
                            font=("Segoe UI", 14, "bold"),
                            width=300, height=40,
                            corner_radius=8,
                            fg_color="#28a745", hover_color="#218838")
        add_btn.pack(side="right")

        # Set focus to username entry
        username_entry.focus_set()

    def edit_user_dialog(self, username):
        """Show enhanced dialog to edit an existing user"""
        edit_win = ctk.CTkToplevel(self.main_app)
        edit_win.transient(self.main_app) # Stay on top
        edit_win.title(f"Edit User: {username}")
        edit_win.attributes("-topmost", True)
        self.center_window(edit_win, 650, 600)
        edit_win.grab_set()

        # Main frame with padding
        main_frame = ctk.CTkFrame(edit_win)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Header with icon and username
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.pack(fill="x", pady=(0, 20))

        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 28)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header_frame, text=f"Edit User",
                    font=("Segoe UI", 22, "bold")).pack(side="left")

        # Username display
        username_frame = ctk.CTkFrame(main_frame, fg_color="#f8f9fa", corner_radius=6)
        username_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(username_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=10)
        ctk.CTkLabel(username_frame, text=username,
                    font=("Segoe UI", 16, "bold")).pack(side="left", padx=5, pady=10)

        # Rest of the function remains the same...
        # Form fields
        form_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        form_frame.pack(fill="x", pady=10)

        # Password field with icon
        password_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        password_frame.pack(fill="x", pady=10)

        ctk.CTkLabel(password_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(password_frame, text=get_translation("ctklabel_text_new_password"),
                    font=("Segoe UI", 14, "bold")).pack(side="left")

        password_entry = ctk.CTkEntry(form_frame, placeholder_text=get_translation("ctkentry_placeholder_text_leave_blank_to_keep_current_password"),
                                    show="•", width=400, height=35, font=("Segoe UI", 14))
        password_entry.pack(pady=(0, 5))

        ctk.CTkLabel(form_frame, text=get_translation("ctklabel_text_note_password_will_only_be_updated_if_a_new_one_is"),
                    font=("Segoe UI", 12),
                    text_color="#6c757d").pack(anchor="w", pady=(0, 15))

        # Show/hide password toggle
        password_visible = False

        def toggle_password_visibility():
            nonlocal password_visible
            password_visible = not password_visible
            password_entry.configure(show="" if password_visible else "•")
            toggle_btn.configure(text="👁️ Hide" if password_visible else "👁️ Show")

        toggle_btn = ctk.CTkButton(form_frame, text=get_translation("ctkbutton_text_show"),
                                command=toggle_password_visibility,
                                font=("Segoe UI", 12),
                                width=80, height=25,
                                fg_color="#6c757d", hover_color="#5a6268")
        toggle_btn.pack(anchor="e", padx=5)

        # Role selection with styled radio buttons
        current_role = users[username]["role"]
        role_frame = ctk.CTkFrame(form_frame, fg_color="transparent")
        role_frame.pack(fill="x", pady=15)

        ctk.CTkLabel(role_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(role_frame, text=get_translation("ctklabel_text_user_role"),
                    font=("Segoe UI", 14, "bold")).pack(side="left")

        role_var = ctk.StringVar(value=current_role)

        role_options = ctk.CTkFrame(form_frame, fg_color="transparent")
        role_options.pack(fill="x", pady=(0, 15))

        # Admin radio with custom styling
        admin_frame = ctk.CTkFrame(role_options, fg_color="#f8f9fa", corner_radius=6)
        admin_frame.pack(side="left", fill="x", expand=True, padx=(0, 5))

        admin_radio = ctk.CTkRadioButton(admin_frame, text=get_translation("ctkradiobutton_text_admin"),
                                        variable=role_var, value="admin",
                                        font=("Segoe UI", 14),
                                        border_width_checked=6,
                                        fg_color="#2D7FF9",
                                        hover_color="#1A6CD6")
        admin_radio.pack(side="left", padx=15, pady=10)

        ctk.CTkLabel(admin_frame, text=get_translation("ctklabel_text_full_system_access"),
                    font=("Segoe UI", 12),
                    text_color="#6c757d").pack(side="left", padx=5)

        # User radio with custom styling
        user_frame = ctk.CTkFrame(role_options, fg_color="#f8f9fa", corner_radius=6)
        user_frame.pack(side="left", fill="x", expand=True, padx=(5, 0))

        user_radio = ctk.CTkRadioButton(user_frame, text=get_translation("ctkradiobutton_text_user"),
                                        variable=role_var, value="user",
                                        font=("Segoe UI", 14),
                                        border_width_checked=6,
                                        fg_color="#28a745",
                                        hover_color="#218838")
        user_radio.pack(side="left", padx=15, pady=10)

        ctk.CTkLabel(user_frame, text=get_translation("ctklabel_text_limited_access"),
                    font=("Segoe UI", 12),
                    text_color="#6c757d").pack(side="left", padx=5)

        # Buttons frame
        buttons_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", pady=(20, 0))

        # Cancel button
        cancel_btn = ctk.CTkButton(buttons_frame, text=get_translation("ctkbutton_text_cancel"),
                                command=edit_win.destroy,
                                font=("Segoe UI", 14),
                                width=120, height=40,
                                fg_color="#6c757d", hover_color="#5a6268")
        cancel_btn.pack(side="left", padx=(0, 10))

        # Inside edit_user_dialog, modify the do_edit_user function:
        def do_edit_user():
            new_password = password_entry.get().strip()
            new_role = role_var.get()

            # Update user details if a new password was entered
            if new_password:
                users[username]["password"] = pbkdf2_sha256.hash(new_password)
            users[username]["role"] = new_role

            # Use the updated password hash when saving
            effective_password = users[username]["password"]
            self.save_user_to_db(username, effective_password, new_role)

            messagebox.showinfo("Success", f"User '{username}' updated successfully", parent=edit_win)
            logging.info(f"Admin '{self.current_user['username']}' updated user '{username}'")
            self.refresh_user_list()
            edit_win.destroy()


        # Update button
        update_btn = ctk.CTkButton(buttons_frame, text=get_translation("ctkbutton_text_update_user"),
                                command=do_edit_user,
                                font=("Segoe UI", 14, "bold"),
                                width=300, height=40,
                                corner_radius=8,
                                fg_color="#007bff", hover_color="#0069d9")
        update_btn.pack(side="right")

    def delete_user_confirm(self, username):
        """Show enhanced confirmation dialog to delete a user"""
        # Don't allow deleting the current user
        if self.current_user and username == self.current_user["username"]:
            messagebox.showerror("Error", "You cannot delete your own account")
            return

            # Create a custom confirmation dialog
                # Create a custom confirmation dialog
        confirm_win = ctk.CTkToplevel(self.main_app)
        confirm_win.transient(self.main_app) # Stay on top
        confirm_win.title("Confirm Delete")
        confirm_win.attributes("-topmost", True)
        self.center_window(confirm_win, 400, 250)
        confirm_win.grab_set()

        # Main frame with padding
        main_frame = ctk.CTkFrame(confirm_win)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Warning icon
        ctk.CTkLabel(main_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 48)).pack(pady=(10, 5))

        # Warning message
        ctk.CTkLabel(main_frame, text=get_translation("ctklabel_text_delete_user"),
                    font=("Segoe UI", 20, "bold")).pack(pady=(0, 10))

        message = f"Are you sure you want to delete user '{username}'?\nThis action cannot be undone."
        ctk.CTkLabel(main_frame, text=message,
                    font=("Segoe UI", 14),
                    justify="center").pack(pady=(0, 20))

        # Buttons frame
        buttons_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", pady=(10, 0))

        # Cancel button
        cancel_btn = ctk.CTkButton(buttons_frame, text=get_translation("ctkbutton_text_cancel"),
                                command=confirm_win.destroy,
                                font=("Segoe UI", 14),
                                width=180, height=40,
                                fg_color="#6c757d", hover_color="#5a6268")
        cancel_btn.pack(side="left")

        def do_delete_user():
            if username in users:
                del users[username]
                # Remove user from the database
                self.cursor.execute("DELETE FROM users WHERE username=?", (username,))
                self.conn.commit()
                messagebox.showinfo("Success", f"User '{username}' deleted successfully", parent=confirm_win)
                logging.info(f"Admin '{self.current_user['username']}' deleted user '{username}'")
                # Refresh the user list
                self.refresh_user_list()
                confirm_win.destroy()
            else:
                messagebox.showerror("Error", f"User '{username}' not found.", parent=confirm_win)

        # Delete button
        delete_btn = ctk.CTkButton(buttons_frame, text=get_translation("ctkbutton_text_delete_user"),
                                command=do_delete_user,
                                font=("Segoe UI", 14, "bold"),
                                width=180, height=40,
                                corner_radius=8,
                                fg_color="#dc3545", hover_color="#c82333")
        delete_btn.pack(side="right")
    def ensure_admin_user_db(self):
        global users
        self.cursor.execute("SELECT COUNT(*) FROM users WHERE role = 'admin'")
        count = self.cursor.fetchone()[0]
        if count == 0:
            # Create default admin user
            admin_hash = pbkdf2_sha256.hash(DEFAULT_ADMIN_PASSWORD)
            self.save_user_to_db("admin", admin_hash, "admin")
            users["admin"] = {"password": admin_hash, "role": "admin"}
            logging.warning("No admin users found, created default admin user.")

    def initialize_user_database(self):
        # --- THIS IS THE KEY CHANGE ---
        data_dir = get_data_dir()
        # ----------------------------

        self.db_path = os.path.join(data_dir, "users.db")
        logging.info(f"Initializing user database at: {self.db_path}")
        self.conn = sqlite3.connect(self.db_path)
        self.cursor = self.conn.cursor()
        # Create table if it doesn't exist
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS users (
                username TEXT PRIMARY KEY,
                password TEXT NOT NULL,
                role TEXT NOT NULL
            )
        """)
        self.conn.commit()
        self.load_users_from_db()
        self.ensure_admin_user_db()

    def save_user_to_db(self, username, hashed_password, role):
        self.cursor.execute("""
            INSERT OR REPLACE INTO users(username, password, role)
            VALUES (?, ?, ?)
        """, (username, hashed_password, role))
        self.conn.commit()
        logging.info(f"User '{username}' saved to database.")


    def load_users_from_db(self):
        global users
        self.cursor.execute("SELECT username, password, role FROM users")
        rows = self.cursor.fetchall()
        users = {username: {"password": password, "role": role} for username, password, role in rows}
        logging.info(f"Loaded {len(users)} users from database.")
        
    def setup_settings_tab(self):
        """Configure the settings tab with improved visuals"""
        # Create a scrollable frame for the settings
        settings_scroll = ctk.CTkScrollableFrame(self.tab_settings)
        settings_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        settings_scroll.columnconfigure(0, weight=1)

        # Appearance Frame with better styling
        appearance_frame = ctk.CTkFrame(settings_scroll, corner_radius=8)
        appearance_frame.pack(fill="x", padx=10, pady=(10, 15))

        # Header with icon
        app_header_frame = ctk.CTkFrame(appearance_frame, fg_color="transparent")
        app_header_frame.pack(fill="x", padx=15, pady=(10, 5))

        ctk.CTkLabel(app_header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(app_header_frame, text=get_translation("ctklabel_text_appearance"),
                    font=("Segoe UI", 18, "bold")).pack(side="left")


        # Theme switcher with improved layout
        theme_frame = ctk.CTkFrame(appearance_frame, fg_color="transparent")
        theme_frame.pack(fill="x", padx=15, pady=15)

        ctk.CTkLabel(theme_frame, text=get_translation("ctklabel_text_select_theme_mode"),
                    font=("Segoe UI", 14)).pack(anchor="w", pady=(0, 10))

        # Radio button group for theme selection
        theme_selection = ctk.CTkFrame(theme_frame, fg_color="transparent")
        theme_selection.pack(fill="x")

        theme_var = ctk.StringVar(value="Dark" if ctk.get_appearance_mode() == "Dark" else "Light")

        def change_theme_mode(choice):
            ctk.set_appearance_mode(choice.lower())
            # Update the status bar notification
            self.notification_label.configure(text=f"Theme changed to {choice} mode")

        # Create radio buttons for theme options
        themes = ["Dark", "Light", "System"]
        theme_radios = []

        for i, theme in enumerate(themes):
            radio = ctk.CTkRadioButton(theme_selection, text=theme, variable=theme_var,
                                    value=theme, command=lambda t=theme: change_theme_mode(t),
                                    font=("Segoe UI", 14))
            radio.pack(side="left", padx=(0 if i == 0 else 20))
            theme_radios.append(radio)

        # Add keyboard shortcut info
        shortcut_frame = ctk.CTkFrame(theme_frame, fg_color="transparent")
        shortcut_frame.pack(fill="x", pady=(10, 0))

        ctk.CTkLabel(shortcut_frame, text=get_translation("ctklabel_text_tip_press_ctrlt_to_quickly_toggle_between_dark_and"),
                    font=("Segoe UI", 12, "italic")).pack(anchor="w")

        # User Settings Frame
        if self.current_user and self.current_user["role"] == "admin":
            user_frame = ctk.CTkFrame(settings_scroll, corner_radius=8)
            user_frame.pack(fill="x", padx=10, pady=15)

            # Header with icon for user settings
            user_header_frame = ctk.CTkFrame(user_frame, fg_color="transparent")
            user_header_frame.pack(fill="x", padx=15, pady=(10, 5))

            ctk.CTkLabel(user_header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(user_header_frame, text=get_translation("ctklabel_text_user_settings"),
                        font=("Segoe UI", 18, "bold")).pack(side="left")
        
            # User info card with shadow effect (simulated with nested frames)
            card_outer = ctk.CTkFrame(user_frame, fg_color=["#D3D3D3", "#2B2B2B"])
            card_outer.pack(fill="x", padx=15, pady=15)

            info_card = ctk.CTkFrame(card_outer)
            info_card.pack(fill="x", padx=2, pady=2)

            # Show user info with icons
            info_frame = ctk.CTkFrame(info_card, fg_color="transparent")
            info_frame.pack(fill="x", padx=15, pady=10)

            # Username row
            username_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
            username_frame.pack(fill="x", pady=5, anchor="w")

            ctk.CTkLabel(username_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(username_frame, text=f"Username: {self.current_user['username']}",
                        font=("Segoe UI", 14, "bold")).pack(side="left")

            # Role row
            role_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
            role_frame.pack(fill="x", pady=5, anchor="w")

            # Icon based on role
            role_icon = "👑" if self.current_user["role"] == "admin" else "🧑"
            ctk.CTkLabel(role_frame, text=role_icon, font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(role_frame, text=f"Role: {self.current_user['role'].capitalize()}",
                        font=("Segoe UI", 14)).pack(side="left")

            # Add last login time (placeholder - in a real app this would be stored)
            login_frame = ctk.CTkFrame(info_frame, fg_color="transparent")
            login_frame.pack(fill="x", pady=5, anchor="w")

            ctk.CTkLabel(login_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 16)).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(login_frame, text=f"Last login: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
                        font=("Segoe UI", 14)).pack(side="left")

            # Account management options
            action_frame = ctk.CTkFrame(user_frame, fg_color="transparent")
            action_frame.pack(fill="x", padx=15, pady=(0, 15))

            # Change password button with icon
            pwd_button_frame = ctk.CTkFrame(action_frame, fg_color="transparent")
            pwd_button_frame.pack(fill="x", pady=5)

            change_pwd_btn = ctk.CTkButton(pwd_button_frame, text=get_translation("ctkbutton_text_change_password"),
                                        command=self.change_password,
                                        font=("Segoe UI", 14),
                                        width=200)
            change_pwd_btn.pack(side="left")
    def setup_manage_tab(self):
        """Configure the manage tab with file management functionality"""
        # Add a scrollable container
        manage_scroll = ctk.CTkScrollableFrame(self.tab_manage)
        manage_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        manage_scroll.grid_columnconfigure(0, weight=1)

        # File Operations Frame with improved styling
        operations_frame = ctk.CTkFrame(manage_scroll, corner_radius=8)
        operations_frame.pack(fill="x", padx=10, pady=(10, 15))

        # Header with icon
        op_header_frame = ctk.CTkFrame(operations_frame, fg_color="transparent")
        op_header_frame.pack(fill="x", padx=15, pady=(10, 5))

        ctk.CTkLabel(op_header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(op_header_frame, text=get_translation("ctklabel_text_file_operations"),
                    font=("Segoe UI", 18, "bold")).pack(side="left")

        # Operations description
        ctk.CTkLabel(operations_frame, text=get_translation("ctklabel_text_manage_your_archived_files_with_these_tools"),
                    font=("Segoe UI", 12)).pack(anchor="w", padx=15, pady=(0, 10))

        # Operations grid with icons and descriptions
        buttons_frame = ctk.CTkFrame(operations_frame, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=15, pady=(5, 15))

        # Define a consistent style for operation buttons
        button_width = 170
        button_height = 80
        button_font = ("Segoe UI", 14, "bold")

        # Create a 2x2 grid of operation buttons
        buttons_frame.grid_columnconfigure((0, 1), weight=1, uniform="equal")

        # Preview button with icon and description
        preview_frame = ctk.CTkFrame(buttons_frame, fg_color="transparent")
        preview_frame.grid(row=0, column=0, padx=5, pady=5, sticky="nsew")

        self.preview_btn = ctk.CTkButton(preview_frame, text=get_translation("ctkbutton_text_preview_file"),
                                        command=self.custom_preview_interface,
                                        font=button_font,
                                        width=button_width, height=button_height,
                                        fg_color=["#3a7ebf", "#1f538d"])
        self.preview_btn.pack(pady=5)

        ctk.CTkLabel(preview_frame, text=get_translation("ctklabel_text_view_image_files"),
                    font=("Segoe UI", 12)).pack()


        # Rollback button
        rollback_frame = ctk.CTkFrame(buttons_frame, fg_color="transparent")
        rollback_frame.grid(row=0, column=1, padx=5, pady=5, sticky="nsew")

        self.rollback_btn = ctk.CTkButton(rollback_frame, text=get_translation("ctkbutton_text_rollback_file"),
                                        command=self.custom_rollback_interface,
                                        font=button_font,
                                        width=button_width, height=button_height,
                                        fg_color=["#3a7ebf", "#1f538d"])
        self.rollback_btn.pack(pady=5)

        ctk.CTkLabel(rollback_frame, text=get_translation("ctklabel_text_restore_previous_versions"),
                    font=("Segoe UI", 12)).pack()

        # Only show activity logs to admin users
        if self.current_user and self.current_user["role"] == "admin":
            # File History and Activity Frame
            history_frame = ctk.CTkFrame(manage_scroll, corner_radius=8)
            history_frame.pack(fill="both", expand=True, padx=10, pady=15)

            # Header with icon
            hist_header_frame = ctk.CTkFrame(history_frame, fg_color="transparent")
            hist_header_frame.pack(fill="x", padx=15, pady=(10, 5))

            ctk.CTkLabel(hist_header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
            ctk.CTkLabel(hist_header_frame, text=get_translation("ctklabel_text_recent_activity"),
                        font=("Segoe UI", 18, "bold")).pack(side="left")

            # Refresh button for activity log
            refresh_btn = ctk.CTkButton(hist_header_frame, text=get_translation("ctkbutton_text_refresh"),
                                        command=self.update_activity_log,
                                        width=100, font=("Segoe UI", 12))
            refresh_btn.pack(side="right", padx=10)

            # Activity log with scrollbar and modern styling
            self.activity_box = ctk.CTkTextbox(history_frame, width=400, height=200,
                                            font=("Consolas", 12))
            self.activity_box.pack(fill="both", expand=True, padx=15, pady=10)
            self.activity_box.configure(state="disabled")

            # Add sample content or load from logs
            self.update_activity_log()
        else:
            # Add a spacer or footer for non-admin users
            footer_frame = ctk.CTkFrame(manage_scroll, corner_radius=8)
            footer_frame.pack(fill="x", padx=10, pady=15)

            ctk.CTkLabel(footer_frame, text=get_translation("ctklabel_text_file_management_tools_are_available_above"),
                        font=("Segoe UI", 14)).pack(pady=20)

    # This method needs to be OUTSIDE of setup_manage_tab (fix the indentation)
    def update_activity_log(self):
        """Update the activity log with recent actions"""
        if not hasattr(self, 'activity_box'):
            return
        self.activity_box.configure(state="normal")
        self.activity_box.delete("1.0", "end")
        try:
            if os.path.exists('archive_app.log'):
                with open('archive_app.log', 'r') as log_file:
                    lines = log_file.readlines()
                    last_lines = lines[-10:] if len(lines) > 10 else lines
                    self.activity_box.insert("1.0", "".join(last_lines))
            else:
                self.activity_box.insert("1.0", "No activity logs found.")
        except Exception as e:
            self.activity_box.insert("1.0", f"Could not load activity logs: {e}")
        self.activity_box.configure(state="disabled")
    def setup_upload_tab(self):
        """Configure the upload tab with file upload functionality,
        including FOUR-level structure and drag & drop."""
        
        # --- Main Scrollable Container ---
        upload_scroll = ctk.CTkScrollableFrame(self.tab_upload)
        upload_scroll.pack(fill="both", expand=True, padx=10, pady=10)
        upload_scroll.grid_columnconfigure(0, weight=1)
        
        # --- Company Selection Frame ---
        company_frame = ctk.CTkFrame(upload_scroll, corner_radius=8)
        company_frame.pack(fill="x", padx=10, pady=(10, 15), anchor="n")
        
        header_frame = ctk.CTkFrame(company_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(header_frame, text=get_translation("ctklabel_text_company_information"), font=("Segoe UI", 18, "bold")).pack(side="left")
        
        company_entry_frame = ctk.CTkFrame(company_frame, fg_color="transparent")
        company_entry_frame.pack(fill="x", padx=15, pady=(5, 15))
        ctk.CTkLabel(company_entry_frame, text=get_translation("ctklabel_text_company_name"), font=("Segoe UI", 14)).pack(side="left", padx=(0, 10))
        self.company_entry = ctk.CTkEntry(company_entry_frame, placeholder_text=get_translation("ctkentry_placeholder_text_enter_company_name"),
                                        font=("Segoe UI", 14))
        self.company_entry.pack(side="left", fill="x", expand=True)
        self.company_entry.bind("<FocusOut>", self.update_options)
        self.company_entry.bind("<Return>", self.update_options)
        # --- Structure Selection Frame ---
        structure_frame = ctk.CTkFrame(upload_scroll, corner_radius=8)
        structure_frame.pack(fill="x", padx=10, pady=15, anchor="n")
        
        struct_header_frame = ctk.CTkFrame(structure_frame, fg_color="transparent")
        struct_header_frame.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(struct_header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(struct_header_frame, text=get_translation("ctklabel_text_document_structure"), font=("Segoe UI", 18, "bold")).pack(side="left")
        ctk.CTkLabel(structure_frame, text=get_translation("ctklabel_text_select_the_appropriate_structure_for_filing"),
                    font=("Segoe UI", 12)).pack(anchor="w", padx=15, pady=(0, 10))
        
        selection_frame = ctk.CTkFrame(structure_frame, fg_color="transparent")
        selection_frame.pack(fill="x", padx=15, pady=(5, 15))
        selection_frame.grid_columnconfigure(1, weight=1)
        
        # Header selection
        ctk.CTkLabel(selection_frame, text=get_translation("ctklabel_text_header"), font=("Segoe UI", 14)).grid(row=0, column=0, sticky="w",
                                                                                padx=(0, 10), pady=10)
        self.header_var = ctk.StringVar(value=list(self.structure.keys())[0])
        self.header_menu = ctk.CTkOptionMenu(selection_frame,
                                            values=list(self.structure.keys()),
                                            variable=self.header_var,
                                            font=("Segoe UI", 13))
        self.header_menu.grid(row=0, column=1, sticky="ew", pady=10)
        self.header_var.trace_add("write", self.update_options)
        
        # Subheader selection
        ctk.CTkLabel(selection_frame, text=get_translation("ctklabel_text_subheader"), font=("Segoe UI", 14)).grid(row=1, column=0, sticky="w",
                                                                                    padx=(0, 10), pady=10)
        self.subheader_var = ctk.StringVar()
        self.subheader_menu = ctk.CTkOptionMenu(selection_frame,
                                                values=[],  # Initialize empty
                                                variable=self.subheader_var,
                                                font=("Segoe UI", 13))
        self.subheader_menu.grid(row=1, column=1, sticky="ew", pady=10)
        self.subheader_var.trace_add("write", self.update_section_options_upload)
        
        # Section selection
        ctk.CTkLabel(selection_frame, text=get_translation("ctklabel_text_section"), font=("Segoe UI", 14)).grid(row=2, column=0, sticky="w",
                                                                                padx=(0, 10), pady=10)
        self.section_var = ctk.StringVar()
        self.section_menu = ctk.CTkOptionMenu(selection_frame,
                                            values=[],  # Initialize empty
                                            variable=self.section_var,
                                            font=("Segoe UI", 13))
        self.section_menu.grid(row=2, column=1, sticky="ew", pady=10)
        # Trace for section change to update subsections
        self.section_var.trace_add("write", self.update_subsection_options_upload)
        
        # --- Subsection Selection (NEW) ---
        ctk.CTkLabel(selection_frame, text=get_translation("ctklabel_text_subsection"), font=("Segoe UI", 14)).grid(row=3, column=0, sticky="w",
                                                                                    padx=(0, 10), pady=10)
        self.subsection_var = ctk.StringVar()
        self.subsection_menu = ctk.CTkOptionMenu(selection_frame,
                                                values=[],  # Initialize empty
                                                variable=self.subsection_var,
                                                font=("Segoe UI", 13))
        self.subsection_menu.grid(row=3, column=1, sticky="ew", pady=10)
        
        # Initialize subheader, section, and subsection options based on initial header selection
        self.update_options()
         # --- << NEW: Add Structure Element Button (Admin Only) >> ---
        if self.current_user and self.current_user["role"] == "admin":
            self.add_folder_button = ctk.CTkButton(
                structure_frame, # Add button to the main structure_frame
                text="➕ Add New Folder Here...",
                font=("Segoe UI", 12), # Slightly smaller font
                height=30,            # Smaller height
                fg_color=("gray70", "gray30"), # Subtle color, adjust as needed
                command=self.add_structure_element_dialog_contextual # Call the NEW dialog function
            )
            # Pack it *after* the selection_frame, within structure_frame
            self.add_folder_button.pack(pady=(10, 15), padx=15, anchor="e") # Add padding, align right
        # --- Upload Options Frame ---
        upload_opts_frame = ctk.CTkFrame(upload_scroll, corner_radius=8)
        upload_opts_frame.pack(fill="x", padx=10, pady=15, anchor="n")
        
        upload_header_frame = ctk.CTkFrame(upload_opts_frame, fg_color="transparent")
        upload_header_frame.pack(fill="x", padx=15, pady=(10, 5))
        ctk.CTkLabel(upload_header_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 20)).pack(side="left", padx=(0, 10))
        ctk.CTkLabel(upload_header_frame, text=get_translation("ctklabel_text_upload_options"), font=("Segoe UI", 18, "bold")).pack(side="left")
        
        file_types_str = ", ".join([ext.upper().replace('.', '') for ext in SUPPORTED_FILE_EXTENSIONS])
        ctk.CTkLabel(upload_opts_frame, text=f"Supported: {file_types_str}",
                    font=("Segoe UI", 12, "italic")).pack(anchor="w", padx=15, pady=(5, 10))
        
        button_frame = ctk.CTkFrame(upload_opts_frame, fg_color="transparent")
        button_frame.pack(fill="x", padx=15, pady=(5, 15))
        
        self.upload_btn = ctk.CTkButton(button_frame, text=get_translation("ctkbutton_text_upload_file"),
                                        command=self.upload_file,
                                        font=("Segoe UI", 14, "bold"),
                                        height=38)
        self.upload_btn.pack(side="left", padx=(0, 10))
        
        self.batch_upload_btn = ctk.CTkButton(button_frame, text=get_translation("ctkbutton_text_batch_upload"),
                                            command=self.batch_upload,
                                            font=("Segoe UI", 14),
                                            height=38)
        self.batch_upload_btn.pack(side="left", padx=10)
        
        self.scan_btn = ctk.CTkButton(button_frame, text=get_translation("ctkbutton_text_scan_archive"),
                                    command=self.scan_and_archive,
                                    font=("Segoe UI", 14, "bold"),
                                    height=38)
        self.scan_btn.pack(side="left", padx=(10, 0))
        
        # --- Drag & Drop Zone ---
        self.dropzone_frame = ctk.CTkFrame(upload_scroll, corner_radius=8, border_width=2,
                                        border_color=("gray70", "gray30"))
        self.dropzone_frame.pack(fill="both", expand=True, padx=10, pady=15, anchor="n")
        
        # Store the original color to restore it later
        self.dropzone_original_color = self.dropzone_frame.cget("border_color")
        
        dropzone_content_frame = ctk.CTkFrame(self.dropzone_frame, fg_color="transparent")
        dropzone_content_frame.pack(expand=True)
        
        self.dropzone_label = ctk.CTkLabel(dropzone_content_frame, text=get_translation("ctklabel_text_drag_drop_files_here"),
                                        font=("Segoe UI", 18, "bold"))
        self.dropzone_label.pack(pady=(5, 5))
        
        upload_icon = ctk.CTkLabel(dropzone_content_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 48))
        upload_icon.pack(pady=10)
        
        self.dropzone_sublabel = ctk.CTkLabel(dropzone_content_frame, text=get_translation("ctklabel_text_or_click_to_browse"),
                                            font=("Segoe UI", 14))
        self.dropzone_sublabel.pack(pady=(5, 5))
        
        clickable_widgets = [
            self.dropzone_frame, dropzone_content_frame,
            self.dropzone_label, upload_icon, self.dropzone_sublabel
        ]
        for widget in clickable_widgets:
            widget.bind("<Button-1>", lambda e: self.upload_file())
            widget.bind("<Enter>", lambda e: widget.configure(cursor="hand2"))
            widget.bind("<Leave>", lambda e: widget.configure(cursor=""))
        
        if DRAG_DROP_ENABLED:
            try:
                dropzone_widget = self.dropzone_frame._canvas if hasattr(self.dropzone_frame, '_canvas') else self.dropzone_frame
                if dropzone_widget:
                    logging.info(f"Registering drop target for widget: {dropzone_widget}")
                    dropzone_widget.drop_target_register(DND_FILES)
                    logging.info("Drop target registered.")
                    dropzone_widget.dnd_bind('<<DropEnter>>', self.on_drop_enter)
                    dropzone_widget.dnd_bind('<<DropLeave>>', self.on_drop_leave)
                    dropzone_widget.dnd_bind('<<Drop>>', self.on_drop)
                    logging.info("Drag & Drop events bound successfully.")
                    self.dropzone_sublabel.configure(text=get_translation("configure_text_or_click_to_browse"))
                else:
                    raise RuntimeError("Could not find underlying Tkinter widget for dropzone.")
            except Exception as e:
                logging.error(f"Error setting up drag and drop: {e}", exc_info=True)
                self.dropzone_sublabel.configure(text="Drag & Drop not available\n(Click to browse)")
        else:
            self.dropzone_sublabel.configure(text="Drag & Drop not available\n(Click to browse)")


    # Inside FileArchiveApp class

    def _perform_folder_creation_task(self, add_struct_win, add_button, original_button_text,
                                      company_display_name, header, current_subheader, current_section,
                                      add_type, new_name):
        """
        Worker thread function to perform the actual folder creation.
        Communicates results back via the UI queue.
        """
        target_level_description = "Unknown" # Default for logging/error
        new_element_path = "Unknown"         # Default for logging/error

        try:
            # --- Ensure Company Structure and Get Safe Name ---
            # This might involve disk operations
            safe_company_name = self.sanitize_path(company_display_name)
            if not hasattr(self, 'current_company') or self.current_company.get("safe_name") != safe_company_name:
                self.create_company_structure(company_display_name)
                safe_company_name = self.current_company["safe_name"]
            logging.debug(f"[_perform_folder_creation_task] Safe company name: {safe_company_name}")

            # --- Determine Structure Type and Target Path ---
            header_structure_type = type(self.structure.get(header))
            is_flat_structure = (header_structure_type is list)
            parent_path = os.path.join(self.archives_path, safe_company_name, header)
            target_level_description = f"Header '{header}'"

            if is_flat_structure:
                if not current_subheader:
                    raise ValueError("No item selected under flat header.") # Error condition
                parent_path = os.path.join(parent_path, current_subheader)
                target_level_description = f"Item '{current_subheader}'"
            else: # Nested structure
                if add_type == "Section":
                    if not current_subheader:
                        raise ValueError("Subheader must be selected to add a new Section.")
                    parent_path = os.path.join(parent_path, current_subheader)
                    target_level_description = f"Subheader '{current_subheader}'"
                elif add_type == "Subsection":
                    if not current_subheader or not current_section:
                        raise ValueError("Subheader and Section must be selected to add a new Subsection.")
                    parent_path = os.path.join(parent_path, current_subheader, current_section)
                    target_level_description = f"Section '{current_section}'"
                else:
                    raise ValueError("Invalid element type selected.")

            # --- Prepare Final Path and Create Directories ---
            new_element_path = os.path.join(parent_path, new_name)
            logging.debug(f"[_perform_folder_creation_task] Target parent path: {parent_path}")
            logging.debug(f"[_perform_folder_creation_task] Target new element path: {new_element_path}")

            # Ensure parent exists (can be slow)
            if not os.path.isdir(parent_path):
                logging.warning(f"[_perform_folder_creation_task] Parent path '{parent_path}' not found. Attempting to create.")
                os.makedirs(parent_path, exist_ok=True)
                logging.info(f"[_perform_folder_creation_task] Created missing parent path: '{parent_path}'")

            # Create target directory (can be slow)
            os.makedirs(new_element_path, exist_ok=True)
            logging.info(f"[_perform_folder_creation_task] Admin '{self.current_user['username']}' created/ensured folder: '{new_element_path}'")

            # --- Success: Queue UI Update ---
            # Prepare data needed for the success callback
            success_info = {
                "new_name": new_name,
                "target_level_description": target_level_description,
                "company_display_name": company_display_name,
                "add_struct_win": add_struct_win, # Pass references carefully
                "add_button": add_button,         # Pass references carefully
                "original_button_text": original_button_text,
                "parent_path_of_new_folder": parent_path # Pass the actual parent path
            }
            # Put the success handling function onto the UI queue
            self.ui_queue.put(lambda info=success_info: self._handle_folder_creation_success(info))

        except (OSError, ValueError, Exception) as e:
            # --- Error: Queue UI Update ---
            logging.error(f"[_perform_folder_creation_task] Failed to create '{new_name}' under '{target_level_description}': {e}", exc_info=True)
            # Prepare data needed for the error callback
            error_info = {
                "error": e,
                "new_name": new_name,
                "add_struct_win": add_struct_win,
                "add_button": add_button,
                "original_button_text": original_button_text
            }
             # Put the error handling function onto the UI queue
            self.ui_queue.put(lambda info=error_info: self._handle_folder_creation_error(info))
    # Inside FileArchiveApp class

    def _handle_folder_creation_success(self, info):
        """Handles UI updates after successful folder creation (called via UI queue)."""
        add_struct_win = info["add_struct_win"]
        add_button = info["add_button"]
        original_button_text = info["original_button_text"]
        new_name = info["new_name"]
        target_level_description = info["target_level_description"]
        company_display_name = info["company_display_name"]
        parent_path_of_new_folder = info.get("parent_path_of_new_folder")

        # Restore UI elements first
        if add_struct_win.winfo_exists():
            add_button.configure(state="normal", text=original_button_text)
            add_struct_win.config(cursor="")
            self.main_app.update() # Ensure UI resets visually

            # Show success message
            messagebox.showinfo("Success",
                                f"Successfully created folder '{new_name}'\n"
                                f"under '{target_level_description}'.\n\n"
                                f"Dropdowns for '{company_display_name}' will refresh automatically.",
                                parent=add_struct_win)

            # Destroy dialog *after* message
            if add_struct_win.winfo_exists():
                 add_struct_win.destroy()

        # Clear cache for the parent path where the new folder was added
        if parent_path_of_new_folder and hasattr(self, 'archive_controller'):
            self.archive_controller.clear_cache(path_prefix=parent_path_of_new_folder)
            logging.info(f"Cleared archive_controller cache for prefix: {parent_path_of_new_folder}")

        # Trigger Dropdown Refresh
        delay_ms = 50 # Shorter delay might be ok now
        logging.info(f"[_handle_folder_creation_success] Scheduling dropdown refresh via update_options() with {delay_ms}ms delay.")
        self.main_app.after(delay_ms, self.update_options)

    def _handle_folder_creation_error(self, info):
        """Handles UI updates after failed folder creation (called via UI queue)."""
        add_struct_win = info["add_struct_win"]
        add_button = info["add_button"]
        original_button_text = info["original_button_text"]
        error = info["error"]
        new_name = info["new_name"]

        # Restore UI elements first
        if add_struct_win.winfo_exists():
            add_button.configure(state="normal", text=original_button_text)
            add_struct_win.config(cursor="")
            self.main_app.update() # Ensure UI resets visually

            # Show specific error message based on caught exception
            if isinstance(error, ValueError):
                 # Validation errors usually show their own message box before raising
                 # So we might not need another one here, or make it more specific.
                 # Let's assume the initial messagebox was sufficient.
                 logging.warning(f"[_handle_folder_creation_error] Handling pre-validated error: {error}")
                 # Optionally show a generic failure message if needed:
                 # messagebox.showerror("Failed", f"Could not add folder due to validation error:\n{error}", parent=add_struct_win)
            elif isinstance(error, OSError):
                 messagebox.showerror("Creation Error", f"Failed to create directory:\n{new_name}\n\nOS Error: {error}", parent=add_struct_win)
            else: # General Exception
                 messagebox.showerror("Unexpected Error", f"An unexpected error occurred while creating '{new_name}':\n{error}", parent=add_struct_win)

            # Don't destroy the dialog on error, allow user to correct input if applicable
            # If you *always* want to close on error, uncomment below:
            # if add_struct_win.winfo_exists():
            #     add_struct_win.destroy()

    # Inside FileArchiveApp class

    def _perform_contextual_folder_add(self, add_struct_win, add_type_var, new_element_entry, add_button):
        """
        Sets loading state and triggers the background task for folder creation.
        Performs only quick, non-blocking validation.
        """
        original_button_text = add_button.cget("text")
        add_struct_win.config(cursor="watch")

        # --- Show Loading State Immediately ---
        try:
            add_button.configure(state="disabled", text=get_translation("configure_text_adding"))
            self.main_app.update() # Force update

            # --- Quick Local Validation ---
            company_display_name = self.company_entry.get().strip()
            header = self.header_var.get() # Needed for context check maybe
            new_name_raw = new_element_entry.get().strip()
            new_name_sanitized = self.sanitize_path(new_name_raw)

            if not company_display_name:
                messagebox.showerror("Error", "Company Name is missing.", parent=add_struct_win)
                raise ValueError("Validation failed: Company Name missing") # Stop before thread

            if not new_name_raw: # Check raw input before sanitizing for emptiness
                messagebox.showerror("Error", "New folder name cannot be empty.", parent=add_struct_win)
                raise ValueError("Validation failed: New folder name empty") # Stop before thread

            # --- Gather Data for Thread ---
            # Get potentially needed context from UI (read-only access is fine here)
            current_subheader = self.subheader_var.get()
            current_section = self.section_var.get()
            add_type = add_type_var.get()

            # --- Launch Background Task ---
            logging.info(f"[Context Add TRIGGER] Launching background task for '{new_name_sanitized}'...")
            self.executor.submit(
                self._perform_folder_creation_task,
                # Pass necessary arguments to the worker thread
                add_struct_win, add_button, original_button_text,
                company_display_name, header, current_subheader, current_section,
                add_type, new_name_sanitized # Pass the sanitized name
            )
            # --- DO NOT RESTORE UI HERE ---
            # UI restoration is now handled by callbacks from the queue

        except ValueError as ve: # Catch local validation errors
             logging.warning(f"[Context Add TRIGGER] Validation failed: {ve}")
             # Restore UI immediately since no thread was started
             if add_struct_win.winfo_exists():
                add_button.configure(state="normal", text=original_button_text)
                add_struct_win.config(cursor="")
                self.main_app.update()
        except Exception as e: # Catch unexpected errors during setup
             logging.error(f"[Context Add TRIGGER] Unexpected error before launching thread: {e}", exc_info=True)
             messagebox.showerror("Error", f"Setup error before adding folder:\n{e}", parent=add_struct_win)
             # Restore UI immediately
             if add_struct_win.winfo_exists():
                 add_button.configure(state="normal", text=original_button_text)
                 add_struct_win.config(cursor="")
                 self.main_app.update()

    def add_structure_element_dialog_contextual(self):
        """Opens a dialog for admins to add a new section or subsection
           using the context (Company, Header, etc.) currently selected
           in the Upload Tab."""

        # --- Get Current Context from Upload Tab ---
        # Needs to happen *before* creating the dialog widgets that depend on it
        try:
            company_display_name = self.company_entry.get().strip()
            header = self.header_var.get()
            subheader = self.subheader_var.get()
            section = self.section_var.get()

            logging.debug(f"[Context Add DIALOG START] Context Read: Co='{company_display_name}', H='{header}', S='{subheader}', Sec='{section}'")

            if not company_display_name or not header:
                messagebox.showerror("Context Error", "Please select a Company and Header in the Upload tab first.", parent=self.main_app)
                return

        except AttributeError as e:
             messagebox.showerror("Error", f"Upload tab elements not fully initialized: {e}", parent=self.main_app)
             logging.error(f"AttributeError getting context: {e}")
             return

        # --- Create the Dialog Window ---
        add_struct_win = ctk.CTkToplevel(self.main_app)
        add_struct_win.transient(self.main_app)
        add_struct_win.title("Add Folder")
        self.center_window(add_struct_win, 550, 350) # Smaller dialog
        add_struct_win.grab_set()
        add_struct_win.attributes("-topmost", True)

        ctk.CTkLabel(add_struct_win, text=get_translation("ctklabel_text_add_new_folder"),
                     font=("Segoe UI", 18, "bold")).pack(pady=(15, 10))

        # (Existing code to display context...)
        context_frame = ctk.CTkFrame(add_struct_win, fg_color="transparent")
        context_frame.pack(fill="x", padx=20, pady=5)
        ctk.CTkLabel(context_frame, text=f"Company: {company_display_name}", font=("Segoe UI", 12)).pack(anchor="w")
        ctk.CTkLabel(context_frame, text=f"Header: {header}", font=("Segoe UI", 12)).pack(anchor="w")
        if subheader:
            ctk.CTkLabel(context_frame, text=f"Subheader: {subheader}", font=("Segoe UI", 12)).pack(anchor="w")
        if section:
            ctk.CTkLabel(context_frame, text=f"Section: {section}", font=("Segoe UI", 12)).pack(anchor="w")

        # (Existing code for type selection...)
        ctk.CTkLabel(add_struct_win, text=get_translation("ctklabel_text_1_what_are_you_adding"), font=("Segoe UI", 14)).pack(anchor="w", padx=20, pady=(15, 2))
        add_type_var = ctk.StringVar(value="Subsection")
        add_type_frame = ctk.CTkFrame(add_struct_win, fg_color="transparent")
        add_type_frame.pack(pady=(0, 10), padx=20)
        ctk.CTkRadioButton(add_type_frame, text=get_translation("ctkradiobutton_text_new_section_under_subheader"), variable=add_type_var, value="Section", font=("Segoe UI", 13)).pack(side="left", padx=10)
        ctk.CTkRadioButton(add_type_frame, text=get_translation("ctkradiobutton_text_new_subsection_under_section"), variable=add_type_var, value="Subsection", font=("Segoe UI", 13)).pack(side="left", padx=10)

        # (Existing code for name entry...)
        ctk.CTkLabel(add_struct_win, text=get_translation("ctklabel_text_2_enter_name_for_new_folder"), font=("Segoe UI", 14)).pack(anchor="w", padx=20, pady=(10, 2))
        new_element_entry = ctk.CTkEntry(add_struct_win, placeholder_text=get_translation("ctkentry_placeholder_text_eg_b10b_or_newsection"),
                                         font=("Segoe UI", 13))
        new_element_entry.pack(pady=(0, 15), padx=20, fill="x")
        new_element_entry.focus_set()

        # --- Action Button (CORRECTED Command Setting) ---
        # Create the button FIRST
        add_button = ctk.CTkButton(add_struct_win, text=get_translation("ctkbutton_text_add_folder"),
                                   # Command will be set below using configure
                                   font=("Segoe UI", 14, "bold"), height=40)
        add_button.pack(pady=(15, 15), padx=20, fill="x")

        # NOW configure the command using lambda to pass the button itself
        add_button.configure(command=lambda: self._perform_contextual_folder_add(
            add_struct_win,      # Argument 1
            add_type_var,        # Argument 2
            new_element_entry,   # Argument 3
            add_button           # Argument 4 <<< THIS WAS MISSING IN THE CALL
        ))
    def scan_and_archive(self):
        """Scan a document via WIA, prompt user for filename with prefix enforcement, and archive it."""
        if not WIA_AVAILABLE:
            messagebox.showerror("Scanning Error", "Scanning via WIA is not available.\nPlease ensure 'comtypes' is installed and you are on Windows.", parent=self.main_app)
            return

        # --- Get Destination Info ---
        company_name = self.company_entry.get().strip()
        if not company_name:
            messagebox.showerror("Error", "Please enter a company name", parent=self.main_app)
            return

        header = self.header_var.get()
        subheader = self.subheader_var.get()
        section = self.section_var.get()
        subsection = self.subsection_var.get() # Get subsection

        try:
            # Ensure structure exists
            self.create_company_structure(company_name)
            safe_company_name = self.current_company["safe_name"]
        except Exception as e:
             messagebox.showerror("Error", f"Failed to create company structure: {e}", parent=self.main_app)
             logging.error(f"[Scan] Error in create_company_structure: {e}")
             return

        # --- Determine Required Prefix and Destination Path ---
        required_prefix = ""
        structure_options = self.structure.get(header, [])
        dest_path = os.path.join(self.archives_path, safe_company_name, header) # Start building path
        has_subsections_defined = False

        if isinstance(structure_options, dict): # Nested Header
            section_dict = structure_options.get(subheader, {})
            if subheader: dest_path = os.path.join(dest_path, subheader) # Add subheader to path

            if section and section in section_dict:
                dest_path = os.path.join(dest_path, section) # Add section path
                subsections_list = section_dict.get(section, [])
                has_subsections_defined = bool(subsections_list)
                if subsection and subsection in subsections_list:
                    required_prefix = subsection
                    dest_path = os.path.join(dest_path, subsection) # Add subsection path
                else:
                    required_prefix = section # Use section prefix if no/invalid subsection
            elif subheader: # Only subheader selected
                 required_prefix = subheader
                 # Path already includes subheader

        elif subheader: # Flat structure, subheader is the final part
            required_prefix = subheader
            dest_path = os.path.join(dest_path, subheader) # Add subheader to path

        # Ensure final directory exists (should be redundant, but safe)
        try:
            os.makedirs(dest_path, exist_ok=True)
        except Exception as e:
            logging.error(f"[Scan] Failed to create destination directory {dest_path}: {e}")
            messagebox.showerror("Error", f"Could not create destination folder:\n{dest_path}\nError: {e}", parent=self.main_app)
            return

        logging.info(f"[Scan] Destination: H='{header}', S='{subheader}', Sec='{section}', SubSec='{subsection}'. Path='{dest_path}'. Required prefix: '{required_prefix}'")

        # --- Perform Scan ---
        self.notification_label.configure(text=get_translation("configure_text_scanning_document_please_wait"))
        self.main_app.update_idletasks()
        scanned_image = None
        try:
            wia = comtypes.client.CreateObject("WIA.CommonDialog")
            # ShowAcquireImage can block, consider running in thread if becomes issue
            scanned_image = wia.ShowAcquireImage()
            if not scanned_image:
                 self.notification_label.configure(text=get_translation("configure_text_scan_cancelled_or_failed"))
                 messagebox.showwarning("Scan Cancelled", "Scan was cancelled or no image was acquired.", parent=self.main_app)
                 return

            self.notification_label.configure(text=get_translation("configure_text_scan_complete_please_name_the_file"))
        except Exception as e:
            logging.error(f"[Scan] Scanning error: {e}", exc_info=True)
            self.notification_label.configure(text=get_translation("configure_text_scan_error"))
            messagebox.showerror("Scan Error", f"Could not scan document: {e}", parent=self.main_app)
            return
        finally:
             # Release COM object if necessary (though comtypes usually handles this)
             # del wia # Maybe not needed
             pass

        # --- Naming Loop ---
        destination_filename = None
        invalid_filename_chars = r'[\\/:"*?<>|]+' # Regex for invalid characters
        # Try to get extension from WIA if possible, otherwise default
        try:
             # This depends on the scanner/WIA version, might not always work
             file_ext = "." + scanned_image.FormatID.split('-')[-1].lower() # Heuristic
             if file_ext not in SUPPORTED_FILE_EXTENSIONS: # Validate guessed extension
                 file_ext = ".png" # Fallback default
        except Exception:
             file_ext = ".png" # Default extension

        while True:
            prompt_title = "Name Scanned File"
            prompt_message = f"Enter a name for the scanned document (extension {file_ext} will be added):"
            base_suggestion = f"scan_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"

            if required_prefix:
                prompt_message = f"Enter name suffix (prefix '{required_prefix}_' will be added automatically):"
                base_suggestion = f"details_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}"

            user_input_part = simpledialog.askstring(
                prompt_title,
                prompt_message,
                parent=self.main_app,
                initialvalue=base_suggestion
            )

            if user_input_part is None: # User cancelled
                messagebox.showinfo("Cancelled", "Scan saving cancelled.", parent=self.main_app)
                self.notification_label.configure(text=get_translation("configure_text_scan_saving_cancelled"))
                logging.info("[Scan] Saving cancelled by user during naming.")
                return # Abort saving

            user_input_part = user_input_part.strip().rstrip('.') # Remove trailing dots/spaces

            if not user_input_part:
                 messagebox.showerror("Invalid Name", "File name part cannot be empty.", parent=self.main_app)
                 continue # Re-prompt

            if re.search(invalid_filename_chars, user_input_part):
                 messagebox.showerror("Invalid Characters", "File name part contains invalid characters (\\ / : * ? \" < > |).", parent=self.main_app)
                 continue # Re-prompt

            # Construct the full name
            if required_prefix:
                # Ensure only one underscore between prefix and user part
                destination_filename = f"{required_prefix.rstrip('_')}_{user_input_part}{file_ext}"
            else:
                destination_filename = f"{user_input_part}{file_ext}"

            logging.info(f"[Scan] User provided name parts, constructed filename: '{destination_filename}'")
            break # Exit the naming loop

        # --- Proceed with Saving ---
        try:
            # Final destination file path
            dest_file = os.path.join(dest_path, destination_filename)
            logging.debug(f"[Scan] Final destination file path: {dest_file}")

            # Handle existing files and backup (using the destination filename)
            if os.path.exists(dest_file):
                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                base, ext = os.path.splitext(destination_filename)
                backup_name = f"{base}_backup_{timestamp}{ext}"
                backup_path = os.path.join(dest_path, backup_name)
                try:
                    count = 0
                    while os.path.exists(backup_path): # Handle collision
                         count += 1
                         backup_name = f"{base}_backup_{timestamp}_{count}{ext}"
                         backup_path = os.path.join(dest_path, backup_name)
                    os.rename(dest_file, backup_path)
                    logging.info(f"[Scan] Existing scan file versioned: {dest_file} -> {backup_path}")
                except (PermissionError, FileNotFoundError, OSError) as e:
                    logging.error(f"[Scan] Error versioning existing scan file {dest_file}: {e}")
                    messagebox.showerror("Save Error", f"Error replacing existing file '{destination_filename}': {e}", parent=self.main_app)
                    return # Stop saving

            # Save the scanned image using the determined destination file path
            scanned_image.SaveFile(dest_file)
            logging.info(f"[Scan] Scanned file saved successfully: {dest_file}")

            self.notification_label.configure(text=get_translation("configure_text_scan_saved_successfully"))
            messagebox.showinfo("Scan Saved", f"Document scanned and saved as:\n'{destination_filename}'\nin folder:\n'{os.path.basename(dest_path)}'", parent=self.main_app) # Show folder name for clarity

        except Exception as e:
            messagebox.showerror("Save Error", f"Failed to save scanned file '{destination_filename}':\n{str(e)}", parent=self.main_app)
            logging.error(f"[Scan] Error saving scan to {dest_path}/{destination_filename}: {str(e)}", exc_info=True)
            self.notification_label.configure(text=get_translation("configure_text_scan_saving_failed"))

    
    def on_drop_enter(self, event):
        """Handle file drag enter event - highlights the drop zone"""
        self.dropzone_frame.configure(border_color=["#1A6CD6", "#2D7FF9"])
        self.dropzone_label.configure(text=get_translation("configure_text_release_to_upload"))

    def on_drop_leave(self, event):
        """Handle file drag leave event - restores the drop zone"""
        self.dropzone_frame.configure(border_color=self.dropzone_original_color)
        self.dropzone_label.configure(text=get_translation("configure_text_drag_drop_files_here"))

    def on_drop(self, event):
        """Handle file drop event - processes dropped files"""
        # Return to normal appearance
        self.dropzone_frame.configure(border_color=self.dropzone_original_color)
        self.dropzone_label.configure(text=get_translation("configure_text_drag_drop_files_here"))

        # Get the dropped file paths
        file_paths = self.parse_drop_data(event.data)

        # Filter for supported files
        valid_files = [f for f in file_paths if self.is_valid_supported_file(f)]

        # Check if any valid files were dropped
        if not valid_files:
            supported_extensions_str = ", ".join([ext.upper().replace('.', '') for ext in SUPPORTED_FILE_EXTENSIONS])
            messagebox.showerror("Invalid Files", f"Please drop only supported files ({supported_extensions_str})")
            return

        # Process the valid files
        self.process_dropped_files(valid_files)

    def parse_drop_data(self, data):
        """Parse the dropped file data into usable file paths"""
        # The data format depends on the platform
        files = []

        if platform.system() == "Windows":
            # Windows format can be complex with curly braces for paths with spaces
            # Example: {C:/path/with spaces/file.txt} {C:/another/file.txt}
            if data.startswith("{") and "}" in data:
                # Extract paths between curly braces
                import re
                files = re.findall(r'{([^}]*)}', data)
            else:
                # If no curly braces, it might be a single file without spaces
                files = [data]
        else:
            # Unix/macOS typically uses newlines
            files = data.split("\n")
            # Remove any empty strings
            files = [f for f in files if f.strip()]

        # Log the parsed files for debugging
        logging.info(f"Parsed dropped files: {files}")

        return files

    def is_valid_supported_file(self, file_path):
        """Check if the file is a valid supported file"""
        _, ext = os.path.splitext(file_path)
        return ext.lower() in SUPPORTED_FILE_EXTENSIONS

    def process_dropped_files(self, file_paths):
        """Process multiple dropped files with naming convention check (Strict Rejection)."""
        company_name = self.company_entry.get().strip()
        if not company_name:
            messagebox.showerror("Error", "Please enter a company name before dropping files.", parent=self.main_app)
            return

        # Get structure selections from the UI
        header = self.header_var.get()
        subheader = self.subheader_var.get()
        section = self.section_var.get()
        subsection = self.subsection_var.get() # Get subsection from main UI state

        # Pre-create company structure once
        try:
            self.create_company_structure(company_name)
        except Exception as e:
             messagebox.showerror("Error", f"Failed to create company structure for drop: {e}", parent=self.main_app)
             logging.error(f"[Drop] Error creating structure before drop: {e}")
             return

        total_files = len(file_paths)
        processed_count = 0
        success_count = 0
        naming_failures = [] # Files rejected due to naming
        other_errors = []
        lock = threading.Lock() # Use lock for shared counters/lists

        self.progress_bar.set(0)
        self.notification_label.configure(text=f"Processing {total_files} dropped files...")
        self.main_app.update_idletasks()

        def drop_task(fp):
            nonlocal processed_count, success_count
            task_success = False
            error_info = None
            intended_drop_filename = os.path.basename(fp) # Start with original

            try:
                # --- Determine required prefix (same logic) ---
                required_prefix = ""
                structure_options_drop = self.structure.get(header, [])
                # ... (logic to determine required_prefix based on header, subheader, section, subsection) ...
                if isinstance(structure_options_drop, dict): # Nested Header
                    section_dict_drop = structure_options_drop.get(subheader, {})
                    if section and section in section_dict_drop:
                        subsections_list_drop = section_dict_drop.get(section, [])
                        if subsection and subsection in subsections_list_drop:
                             required_prefix = subsection
                        else:
                            required_prefix = section
                    elif subheader:
                        required_prefix = subheader
                elif subheader: # Flat structure
                     required_prefix = subheader


                # --- Check naming for drop (NO auto-rename) ---
                needs_rename_drop = False
                if required_prefix:
                    if not intended_drop_filename.startswith(required_prefix + "_") and not intended_drop_filename.startswith(required_prefix):
                         needs_rename_drop = True

                if needs_rename_drop:
                    # Strict check for Drag & Drop - add to failures and return
                    logging.warning(f"[Drop Task] Rejecting {fp}: Prefix mismatch ('{required_prefix}' needed).")
                    with lock:
                        naming_failures.append(os.path.basename(fp))
                    return # Exit this task

                # --- Call perform_file_upload with the determined final name (which is original here) ---
                result = self.perform_file_upload(
                    company_name, header, subheader, section, subsection, # Structure
                    fp,                      # Source path
                    intended_drop_filename   # Final name (original, as rename check passed)
                )
                task_success = result # True or raises Exception

            except Exception as e:
                 error_info = f"{os.path.basename(fp)}: {e}"
                 logging.error(f"[Drop Task] Error during upload for {fp}: {e}", exc_info=True)
                 with lock:
                     other_errors.append(error_info)
            finally:
                # --- (Update counters and UI as before) ---
                with lock:
                    processed_count += 1
                    if task_success:
                        success_count += 1
                # Update progress via queue
                progress = processed_count / total_files
                status_msg = f"Processing drop {processed_count}/{total_files}..."
                self.ui_queue.put(lambda p=progress, msg=status_msg: (
                    self.progress_bar.set(p),
                    self.notification_label.configure(text=msg)
                ))

        # Submit drop tasks
        futures = [self.executor.submit(drop_task, fp) for fp in file_paths]

        # Monitor completion (similar to batch upload)
        def check_drop_completion():
            all_done = all(f.done() for f in futures)
            if all_done:
                logging.info(f"[Drop] All {total_files} tasks completed. Success: {success_count}, Naming Rejected: {len(naming_failures)}, Errors: {len(other_errors)}")
                # Schedule the final report via UI queue
                self.ui_queue.put(lambda: self.report_batch_results(total_files, success_count, naming_failures, other_errors)) # Reuse report function
            else:
                self.main_app.after(200, check_drop_completion)

        self.main_app.after(200, check_drop_completion)
    # --------------------------------------------------------------------------
    # Company Structure & File Upload
    # --------------------------------------------------------------------------
    def upload_file(self):
        """
        Handles single file upload via dialog, including interactive renaming check
        based on Header/Subheader/Section/Subsection selection.
        """
        company_name = self.company_entry.get().strip()
        if not company_name:
            messagebox.showerror("Error", "Please enter a company name", parent=self.main_app)
            return

        header = self.header_var.get()
        subheader = self.subheader_var.get()
        section = self.section_var.get()
        subsection = self.subsection_var.get() # Get subsection value

        # Verify selections are valid (prevent uploading to blank selections where not allowed)
        # Add checks if needed, e.g., ensure section is selected if structure requires it

        try:
            # Ensure structure exists *before* file dialog
            self.create_company_structure(company_name)
            safe_company_name = self.current_company["safe_name"]
        except Exception as e:
             messagebox.showerror("Error", f"Failed to create company structure: {e}", parent=self.main_app)
             logging.error(f"Error in create_company_structure before upload: {e}")
             return

        file_path = filedialog.askopenfilename(
            title="Select a file to upload",
            parent=self.main_app, # Make dialog modal to app
            filetypes=(
                ("Supported files", " ".join([f"*{ext}" for ext in SUPPORTED_FILE_EXTENSIONS])),
                ("All files", "*.*")
            )
        )

        if not file_path:
            self.notification_label.configure(text=get_translation("configure_text_file_selection_cancelled"))
            return

        original_filename = os.path.basename(file_path)
        destination_filename = original_filename # Start with the original name

        # --- Determine Required Prefix (same logic as perform_file_upload) ---
        required_prefix = ""
        structure_options = self.structure.get(header, [])
        has_subsections_defined = False

        if isinstance(structure_options, dict): # Nested Header
            section_dict = structure_options.get(subheader, {})
            if section and section in section_dict:
                subsections_list = section_dict.get(section, [])
                has_subsections_defined = bool(subsections_list)
                if subsection and subsection in subsections_list:
                     required_prefix = subsection
                else:
                    required_prefix = section
            elif subheader:
                required_prefix = subheader
        elif subheader: # Flat structure
             required_prefix = subheader

        logging.info(f"[UploadSingle] Prefix Check: H='{header}', S='{subheader}', Sec='{section}', SubSec='{subsection}'. Required: '{required_prefix}' for '{original_filename}'")

        # --- Enforce Naming Convention Interactively ---
        # Check if prefix exists and if filename *doesn't* start with prefix + "_"
        # Or if it doesn't start exactly with the prefix (for cases like 'c1')
        needs_rename = False
        if required_prefix:
             if not original_filename.startswith(required_prefix + "_") and not original_filename.startswith(required_prefix):
                  needs_rename = True
             elif original_filename.startswith(required_prefix + "_"):
                  logging.debug(f"[UploadSingle] File '{original_filename}' already has correct prefix '{required_prefix}_'.")
             elif original_filename.startswith(required_prefix) and not original_filename.startswith(required_prefix + "_"):
                  # This case might be okay if prefix is standalone like 'c1'.
                  # Or you might want to enforce 'c1_...' ? Let's allow exact match for now.
                  logging.debug(f"[UploadSingle] File '{original_filename}' starts exactly with prefix '{required_prefix}'. Allowing.")


        if needs_rename:
            logging.warning(f"[UploadSingle] File '{original_filename}' needs prefix '{required_prefix}'. Prompting user.")

            proposed_name = f"{required_prefix}_{original_filename}"
            confirm_auto = messagebox.askyesno(
                "Automatic Rename?",
                f"File name '{original_filename}' should start with '{required_prefix}_'.\n\n"
                f"Automatically rename it to:\n'{proposed_name}'?",
                parent=self.main_app
            )

            if confirm_auto: # User clicked YES
                destination_filename = proposed_name
                logging.info(f"[UploadSingle] User confirmed auto-rename to '{destination_filename}'")
            else: # User clicked NO - Manual Rename Loop
                logging.info("[UploadSingle] User declined auto-rename. Prompting for manual name.")
                while True:
                    new_name_suggestion = proposed_name
                    new_name = simpledialog.askstring(
                        "Manual Rename Required",
                        f"File name must start with '{required_prefix}_'.\n"
                        f"Current name: '{original_filename}'\n\n"
                        f"Enter the new name (must start with '{required_prefix}_'):",
                        parent=self.main_app,
                        initialvalue=new_name_suggestion
                    )

                    if new_name is None: # User pressed Cancel
                        messagebox.showinfo("Cancelled", "Upload cancelled during manual renaming.", parent=self.main_app)
                        self.notification_label.configure(text=get_translation("configure_text_upload_cancelled"))
                        logging.info("[UploadSingle] Upload cancelled by user during manual rename.")
                        return # Abort upload

                    new_name = new_name.strip()
                    if not new_name:
                         messagebox.showerror("Invalid Name", "New file name cannot be empty.", parent=self.main_app)
                         continue # Re-prompt manual loop
                    if not new_name.startswith(required_prefix + "_"):
                        messagebox.showerror("Invalid Prefix", f"The new name MUST start with '{required_prefix}_'.", parent=self.main_app)
                        continue # Re-prompt manual loop

                    # Valid manual name entered
                    destination_filename = new_name
                    logging.info(f"[UploadSingle] User provided valid manual name: '{destination_filename}'")
                    break # Exit the manual renaming loop

        # --- Proceed with Upload using destination_filename ---
        try:
            self.notification_label.configure(text=f"Uploading {destination_filename}...")
            self.main_app.update_idletasks() # Show status update

            # --- MODIFIED CALL ---
            # Call perform_file_upload, passing the final filename determined above
            success = self.perform_file_upload(
                company_name, header, subheader, section, subsection, # Structure info
                file_path,                     # The original source file path
                destination_filename           # The final filename after potential rename
            )
            # --- END MODIFIED CALL ---

            if success: # perform_file_upload now only returns True or raises Exception
                # Recalculate final path for display message (as before)
                final_dest_path = os.path.join(self.archives_path, safe_company_name, header)
                # ... (rest of path calculation logic for message box) ...
                structure_options_disp = self.structure.get(header, [])
                has_subsections_defined_disp = False
                if isinstance(structure_options_disp, dict):
                    if subheader: final_dest_path = os.path.join(final_dest_path, subheader)
                    section_dict_disp = structure_options_disp.get(subheader, {})
                    if section:
                        final_dest_path = os.path.join(final_dest_path, section)
                        subsections_list_disp = section_dict_disp.get(section, [])
                        has_subsections_defined_disp = bool(subsections_list_disp)
                        if subsection and has_subsections_defined_disp:
                             final_dest_path = os.path.join(final_dest_path, subsection)
                elif subheader:
                    final_dest_path = os.path.join(final_dest_path, subheader)

                self.notification_label.configure(text=f"Uploaded: {destination_filename}")
                messagebox.showinfo("Success", f"File uploaded successfully as:\n'{destination_filename}'\nto:\n{final_dest_path}", parent=self.main_app)
                logging.info(f"[UploadSingle] Complete: {os.path.join(final_dest_path, destination_filename)} for company {company_name}")
            # else: # This 'else' case is no longer reachable if perform_file_upload only returns True or raises Exception
            #     logging.error("[UploadSingle] Upload failed unexpectedly (perform_file_upload returned False).")
            #     self.notification_label.configure(text="Upload failed (internal error).")
            #     messagebox.showerror("Error", "Upload failed due to an internal error.", parent=self.main_app)

        except Exception as e: # Catch exceptions raised by perform_file_upload
            # Log the full exception
            logging.error(f"[UploadSingle] Upload error for {file_path} (intended name {destination_filename}): {str(e)}", exc_info=True)
            # Provide specific feedback
            if isinstance(e, IOError):
                 err_msg = f"Failed to save file '{destination_filename}':\n{str(e)}\n\nCheck permissions and disk space."
            else:
                 err_msg = f"An unexpected error occurred during upload:\n{str(e)}"

            messagebox.showerror("Upload Error", err_msg, parent=self.main_app)
            self.notification_label.configure(text=get_translation("configure_text_upload_failed"))

    def batch_upload(self):
        """Handles batch file upload with pre-check for naming and optional auto-rename."""
        company_name = self.company_entry.get().strip()
        header = self.header_var.get()
        subheader = self.subheader_var.get()
        section = self.section_var.get()
        subsection = self.subsection_var.get() # Get subsection

        if not company_name:
            messagebox.showerror("Error", "Please enter a company name.", parent=self.main_app)
            return
        # Optional: Add validation for header/subheader/section/subsection if needed

        try:
            # Pre-create company structure once before the loop
            self.create_company_structure(company_name)
            safe_company_name = self.current_company["safe_name"] # Get safe name
        except Exception as e:
             messagebox.showerror("Error", f"Failed to create company structure for batch: {e}", parent=self.main_app)
             logging.error(f"[Batch] Error creating structure before batch: {e}")
             return

        file_paths = filedialog.askopenfilenames(
            title="Select files for batch upload",
            parent=self.main_app,
            filetypes=[("Supported Files", " ".join([f"*{ext}" for ext in SUPPORTED_FILE_EXTENSIONS]))]
        )
        if not file_paths:
            self.notification_label.configure(text=get_translation("configure_text_batch_upload_cancelled"))
            return

        total_files = len(file_paths)
        files_needing_rename = []
        determined_prefix = "" # Store the prefix determined during the check

        # --- Pre-check for Naming Convention ---
        logging.info(f"[Batch] Starting pre-check for {total_files} files...")
        structure_options = self.structure.get(header, []) # Get structure once

        for fp in file_paths:
            original_filename = os.path.basename(fp)

            # Determine required prefix (same logic as perform_file_upload)
            required_prefix = ""
            has_subsections_defined = False
            if isinstance(structure_options, dict): # Nested Header
                section_dict = structure_options.get(subheader, {})
                if section and section in section_dict:
                    subsections_list = section_dict.get(section, [])
                    has_subsections_defined = bool(subsections_list)
                    if subsection and subsection in subsections_list:
                         required_prefix = subsection
                    else:
                        required_prefix = section
                elif subheader:
                    required_prefix = subheader
            elif subheader: # Flat structure
                 required_prefix = subheader

            if not determined_prefix and required_prefix: # Store first determined prefix
                determined_prefix = required_prefix

            # Check naming convention (Prefix_ or exact Prefix)
            needs_rename = False
            if required_prefix:
                if not original_filename.startswith(required_prefix + "_") and not original_filename.startswith(required_prefix):
                     needs_rename = True

            if needs_rename:
                files_needing_rename.append(original_filename)
                logging.debug(f"[Batch PreCheck] File '{original_filename}' needs prefix '{required_prefix}'.")


        # --- Ask for Confirmation if Renaming is Needed ---
        auto_rename_confirmed = False # Default to false (skip misnamed files)
        if files_needing_rename:
            num_to_rename = len(files_needing_rename)
            # Use askyesnocancel: Yes=Rename, No=Skip, Cancel=Abort
            response = messagebox.askyesnocancel(
                "Batch Rename Confirmation",
                f"{num_to_rename} selected file(s) do not start with the required prefix ('{determined_prefix}').\n\n"
                f"- YES: Automatically rename these {num_to_rename} file(s) with the prefix and upload.\n"
                f"- NO: Upload only files that already have the correct name (SKIP the {num_to_rename}).\n"
                f"- CANCEL: Abort the entire batch upload.",
                icon='warning', # Add an icon
                parent=self.main_app # Ensure dialog is modal to the app
            )

            if response is True: # User clicked YES
                auto_rename_confirmed = True
                logging.info("[Batch] User confirmed automatic renaming for batch.")
            elif response is False: # User clicked NO
                auto_rename_confirmed = False
                logging.info("[Batch] User declined automatic renaming. Incorrectly named files will be skipped.")
            else: # User clicked Cancel (response is None)
                messagebox.showinfo("Batch Cancelled", "Batch upload cancelled by user.", parent=self.main_app)
                self.notification_label.configure(text=get_translation("configure_text_batch_upload_cancelled"))
                logging.info("[Batch] Upload cancelled by user during rename confirmation.")
                return # Abort the batch operation


        # --- Proceed with Threaded Upload ---
        self.progress_bar.set(0)
        self.notification_label.configure(text=f"Starting batch upload of {total_files} files...")
        self.main_app.update_idletasks()

        processed_count = 0
        success_count = 0
        naming_failures = [] # Files skipped because naming failed AND auto_rename was False
        other_errors = []
        lock = threading.Lock()

        def upload_task(fp):
            nonlocal processed_count, success_count
            task_success = False
            error_info = None
            intended_batch_filename = os.path.basename(fp) # Start with original

            try:
                # --- Determine required prefix (same logic as before) ---
                required_prefix = ""
                structure_options_task = self.structure.get(header, [])
                # ... (logic to determine required_prefix based on header, subheader, section, subsection) ...
                if isinstance(structure_options_task, dict): # Nested Header
                    section_dict_task = structure_options_task.get(subheader, {})
                    if section and section in section_dict_task:
                        subsections_list_task = section_dict_task.get(section, [])
                        # has_subsections_defined_task = bool(subsections_list_task) # Not strictly needed here
                        if subsection and subsection in subsections_list_task:
                             required_prefix = subsection
                        else:
                            required_prefix = section
                    elif subheader:
                        required_prefix = subheader
                elif subheader: # Flat structure
                     required_prefix = subheader

                # --- Determine final filename based on auto_rename_confirmed ---
                needs_rename_batch = False
                if required_prefix:
                    if not intended_batch_filename.startswith(required_prefix + "_") and not intended_batch_filename.startswith(required_prefix):
                         needs_rename_batch = True

                if needs_rename_batch:
                    if auto_rename_confirmed: # Flag set based on user dialog earlier
                        intended_batch_filename = f"{required_prefix}_{intended_batch_filename}"
                        logging.info(f"[Batch Task] Auto-renaming to: {intended_batch_filename}")
                    else:
                        # Skip this file - add to naming failures and return
                        logging.warning(f"[Batch Task] Skipping {fp}: Prefix mismatch and auto-rename declined.")
                        with lock:
                            naming_failures.append(os.path.basename(fp))
                        # Explicitly return here to skip calling perform_file_upload for this file
                        return # Exit this task for this file

                # --- Call perform_file_upload with the determined final name ---
                result = self.perform_file_upload(
                    company_name, header, subheader, section, subsection, # Structure
                    fp,                      # Source path
                    intended_batch_filename  # Final name for archive
                )
                # perform_file_upload now returns True or raises Exception
                task_success = result # Should be True if no exception raised

            except Exception as e:
                error_info = f"{os.path.basename(fp)} -> {intended_batch_filename}: {e}"
                logging.error(f"[Batch Task] Error during upload for {fp}: {e}", exc_info=True)
                with lock:
                    other_errors.append(error_info)
            finally:
                # --- (Update counters and UI as before) ---
                with lock:
                    processed_count += 1
                    if task_success:
                        success_count += 1
                # Update progress via queue for thread safety
                progress = processed_count / total_files
                status_msg = f"Processing {processed_count}/{total_files}..."
                self.ui_queue.put(lambda p=progress, msg=status_msg: (
                    self.progress_bar.set(p),
                    self.notification_label.configure(text=msg)
                ))
        # --- END OF UPLOAD TASK ---
        # Submit tasks to the shared executor
        futures = [self.executor.submit(upload_task, fp) for fp in file_paths]

        # Monitor completion using 'after' to avoid blocking UI
        def check_completion():
            # Check if all futures are done
            all_done = all(f.done() for f in futures)
            if all_done:
                # All tasks finished, report results via UI queue
                logging.info(f"[Batch] All {total_files} tasks completed. Success: {success_count}, Naming Skipped: {len(naming_failures)}, Errors: {len(other_errors)}")
                # Check for exceptions in futures (optional but good practice)
                for f in futures:
                    if f.exception():
                        # Exceptions should have been caught in upload_task, but log if any leaked
                        logging.error(f"[Batch] Future reported an exception (should have been caught): {f.exception()}")
                # Schedule the final report via UI queue
                self.ui_queue.put(lambda: self.report_batch_results(total_files, success_count, naming_failures, other_errors))
            else:
                # Not all done, schedule another check
                self.main_app.after(200, check_completion) # Check again in 200ms

        # Start the first check
        self.main_app.after(200, check_completion)


    def report_batch_results(self, total, success, naming_fails, other_errs):
        """Updates UI after batch upload completion."""
        message_lines = [f"Batch Upload Report ({success}/{total} successful):"]
        status_text = f"Batch complete: {success}/{total} succeeded."

        if naming_fails:
            fail_list_preview = ", ".join(naming_fails[:3]) # Show first few
            suffix = f" and {len(naming_fails) - 3} more" if len(naming_fails) > 3 else ""
            message_lines.append(f"\nNaming/Skipped ({len(naming_fails)}):")
            message_lines.extend([f"- {nf}" for nf in naming_fails]) # Add full list to details
            status_text += f" {len(naming_fails)} skipped (naming)."
            logging.warning(f"[Batch Report] Naming failures/skipped: {naming_fails}")

        if other_errs:
            error_list_preview = ", ".join([e.split(':')[0] for e in other_errs[:3]]) # Show filenames from errors
            suffix = f" and {len(other_errs) - 3} more" if len(other_errs) > 3 else ""
            message_lines.append(f"\nOther Errors ({len(other_errs)}):")
            message_lines.extend([f"- {oe}" for oe in other_errs]) # Add full list to details
            status_text += f" {len(other_errs)} errors."
            logging.error(f"[Batch Report] Other errors: {other_errs}")

        # Show summary message box
        if naming_fails or other_errs:
             # Show a more detailed message box if there were issues
             detailed_message = "\n".join(message_lines)
             # Limit message box height/content if necessary
             max_len = 1500 # Limit message length
             if len(detailed_message) > max_len:
                 detailed_message = detailed_message[:max_len] + "\n\n... (See log for full details)"
             messagebox.showwarning("Batch Upload Issues", detailed_message, parent=self.main_app)
        else:
            messagebox.showinfo("Batch Upload Complete", f"Successfully uploaded {success} of {total} files.", parent=self.main_app)

        # Update status bar
        self.notification_label.configure(text=status_text)
        # Reset progress bar after a delay
        self.main_app.after(5000, lambda: self.progress_bar.set(0)) # Longer delay
        # Optionally, reset status bar text after even longer
        self.main_app.after(10000, lambda: self.notification_label.configure(text=get_translation("configure_text_ready")))

    # --------------------------------------------------------------------------
    # Dashboard and Search Functionality
    # --------------------------------------------------------------------------
    def update_search_results(self, results):
        """Update the search results UI from background scan."""
        for widget in self.results_frame.winfo_children():
            widget.destroy()
        if results:
            for full_path, name, mod_time in results:
                if os.path.isfile(full_path):
                    try:
                        size = os.stat(full_path).st_size
                        mod_time_str = mod_time.strftime("%Y-%m-%d %H:%M:%S") if mod_time else "Unknown"
                    except Exception:
                        size, mod_time_str = "Unknown", "Unknown"
                    details = f"File: {name}\nPath: {full_path}\nSize: {size} bytes | Modified: {mod_time_str}"
                else:
                    details = f"Folder: {name}\nPath: {full_path}"
                btn = ctk.CTkButton(self.results_frame, text=details, anchor="w",
                                    command=lambda p=full_path: self.open_path(p), font=("Segoe UI", 12))
                btn.pack(pady=2, fill="x", padx=5)
        else:
            ctk.CTkLabel(self.results_frame, text=get_translation("ctklabel_text_no_results_found"), font=("Segoe UI", 12)).pack(pady=5)
    # --- In search_archive method (replace its inner function perform_search) ---
    def search_archive(self):
        """Search the archive with enhanced UI—run heavy scanning off the UI thread."""
        self.temp_archive_path = tempfile.mkdtemp()
        logging.info(f"Created temporary directory for archive search: {self.temp_archive_path}")

        def perform_search():
            query = self.search_entry.get().strip().lower()
            if query:
                with self.search_queries_lock:
                    self.search_queries.append(query)
            file_type_filter = self.file_type_var.get()  # "All" or "Images"
            try:
                start_date = datetime.datetime.strptime(self.start_date_entry.get().strip(), "%Y-%m-%d") \
                    if self.start_date_entry.get().strip() else None
            except ValueError:
                self.ui_queue.put(lambda: messagebox.showerror("Error", "Start date format should be YYYY-MM-DD"))
                return
            try:
                end_date = datetime.datetime.strptime(self.end_date_entry.get().strip(), "%Y-%m-%d") \
                    if self.end_date_entry.get().strip() else None
            except ValueError:
                self.ui_queue.put(lambda: messagebox.showerror("Error", "End date format should be YYYY-MM-DD"))
                return

            results = []
            for root, dirs, files in os.walk(self.archives_path):
                for name in dirs + files:
                    if query in name.lower():
                        full_path = os.path.join(root, name)
                        if file_type_filter != "All" and os.path.isfile(full_path):
                            ext = os.path.splitext(name)[1].lower()
                            if file_type_filter == "Images" and ext not in SUPPORTED_FILE_EXTENSIONS:
                                continue
                        try:
                            mod_time = datetime.datetime.fromtimestamp(os.stat(full_path).st_mtime)
                        except Exception:
                            mod_time = None
                        if start_date and mod_time and mod_time < start_date:
                            continue
                        if end_date and mod_time and mod_time > end_date:
                            continue
                        results.append((full_path, name, mod_time))
            self.ui_queue.put(lambda: self.update_search_results(results))
            logging.info(f"Search for '{query}' returned {len(results)} results.")

        # Run scanning in background:
        self.executor.submit(perform_search)

        def open_search_window():
            search_win = ctk.CTkToplevel(self.main_app)
            search_win.transient(self.main_app) # Stay on top
            search_win.title("Search Archive")
            self.center_window(search_win, 700, 600)
            search_win.grab_set()

            ctk.CTkLabel(search_win, text=get_translation("ctklabel_text_search_archive"), font=("Segoe UI", 18, "bold")).pack(pady=10)
            search_frame = ctk.CTkFrame(search_win)
            search_frame.pack(pady=5, padx=5, fill="x")
            ctk.CTkLabel(search_frame, text=get_translation("ctklabel_text_search_query"), font=("Segoe UI", 14)).grid(row=0, column=0, padx=5, pady=5, sticky="w")
            self.search_entry = ctk.CTkEntry(search_frame, placeholder_text=get_translation("ctkentry_placeholder_text_enter_search_term"), width=200, font=("Segoe UI", 14))
            self.search_entry.grid(row=0, column=1, padx=5, pady=5)
            ctk.CTkLabel(search_frame, text=get_translation("ctklabel_text_file_type"), font=("Segoe UI", 14)).grid(row=1, column=0, padx=5, pady=5, sticky="w")
            self.file_type_var = ctk.StringVar(value="All")
            file_type_menu = ctk.CTkOptionMenu(search_frame, variable=self.file_type_var, values=["All", "Images"], font=("Segoe UI", 14))
            file_type_menu.grid(row=1, column=1, padx=5, pady=5)
            ctk.CTkLabel(search_frame, text=get_translation("ctklabel_text_start_date_yyyymmdd"), font=("Segoe UI", 14)).grid(row=2, column=0, padx=5, pady=5, sticky="w")
            self.start_date_entry = ctk.CTkEntry(search_frame, placeholder_text=get_translation("ctkentry_placeholder_text_yyyymmdd"), width=200, font=("Segoe UI", 14))
            self.start_date_entry.grid(row=2, column=1, padx=5, pady=5)
            ctk.CTkLabel(search_frame, text=get_translation("ctklabel_text_end_date_yyyymmdd"), font=("Segoe UI", 14)).grid(row=3, column=0, padx=5, pady=5, sticky="w")
            self.end_date_entry = ctk.CTkEntry(search_frame, placeholder_text=get_translation("ctkentry_placeholder_text_yyyymmdd"), width=200, font=("Segoe UI", 14))
            self.end_date_entry.grid(row=3, column=1, padx=5, pady=5)
            ctk.CTkButton(search_frame, text=get_translation("ctkbutton_text_search"), command=perform_search, font=("Segoe UI", 14)).grid(row=4, column=0, columnspan=2, pady=10)
            self.results_frame = ctk.CTkScrollableFrame(search_win, width=680, height=350)
            self.results_frame.pack(pady=10, padx=10)
        open_search_window()

    # --------------------------------------------------------------------------
    # Open Path (using OS commands, for admin use if needed)
    # --------------------------------------------------------------------------
    def open_path(self, path):
        try:
            if platform.system() == "Windows":
                os.startfile(path)
            elif platform.system() == "Darwin":
                subprocess.call(["open", path])
            else:
                subprocess.call(["xdg-open", path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open {path}: {e}")
    def logout(self):
        """Log out the current user and return to login screen"""
        if not self.current_user:
            return

        # Hide the archive folder
        self.hide_archive_folder()
        logging.info("Archive folder hidden on logout")

        # Log the logout
        logging.info(f"User '{self.current_user['username']}' logged out")

        # Clear current user
        self.current_user = None

        # Hide the main window
        self.main_app.withdraw()

        # Show the login dialog
        self.authenticate_user()

    # --------------------------------------------------------------------------
    # Company Structure & File Upload
    # --------------------------------------------------------------------------
    def sanitize_path(self, text):
        """Sanitize text for use in file paths, preserving Unicode characters"""
        # Replace filesystem-unsafe characters with safe ones
        invalid_chars = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        for char in invalid_chars:
            text = text.replace(char, '_')

        # Ensure the path is valid for the current OS
        if platform.system() == 'Windows':
            # Handle Windows-specific path limitations (260 char path limit, trailing spaces/dots)
            text = text.rstrip('. ')
            # Prefix long paths with \\?\ on Windows to handle paths > 260 chars
            # (only needed for very long paths, but good practice)
            if len(text) > 200:  # Conservative threshold
                return f"\\\\?\\{text}"

        return text
    def create_company_structure(self, company_name):
        """Create folder structure for a company with Unicode support, including subsections."""
        try:
            safe_company_name = self.sanitize_path(company_name)
            # Update current company info
            self.current_company = {"display_name": company_name, "safe_name": safe_company_name}

            base_path = os.path.join(self.archives_path, safe_company_name)
            logging.info(f"Creating/Verifying structure for company: {company_name} (Safe Path: {safe_company_name})")

            # Create base company folder
            logging.debug(f"Checking/Creating Company Base: {base_path}")
            os.makedirs(base_path, exist_ok=True)

            # Iterate through defined structure
            for header, sub_level_data in self.structure.items():
                header_path = os.path.join(base_path, header)
                logging.debug(f"  Checking/Creating Header: {header_path}")
                os.makedirs(header_path, exist_ok=True)

                if isinstance(sub_level_data, dict): # Nested: Header -> Subheader -> Section -> Subsection
                    logging.debug(f"    Processing Nested Header: {header}")
                    for subheader, section_dict in sub_level_data.items():
                        subheader_path = os.path.join(header_path, subheader)
                        logging.debug(f"      Checking/Creating Subheader: {subheader_path}")
                        os.makedirs(subheader_path, exist_ok=True)

                        if isinstance(section_dict, dict): # Ensure it's the expected Section dict
                            logging.debug(f"        Processing Section Dictionary for Subheader: {subheader}")
                            for section, subsection_list in section_dict.items():
                                section_path = os.path.join(subheader_path, section)
                                logging.debug(f"          Checking/Creating Section: {section_path}")
                                os.makedirs(section_path, exist_ok=True)

                                # Create subsection folders if defined
                                if isinstance(subsection_list, list):
                                    if subsection_list: # Only log if there are subsections to create
                                        logging.debug(f"            Processing Subsection List for Section: {section}")
                                    for subsection in subsection_list:
                                        if subsection: # Skip if subsection name is empty string or None
                                             subsection_path = os.path.join(section_path, subsection)
                                             logging.debug(f"              Checking/Creating Subsection: {subsection_path}")
                                             os.makedirs(subsection_path, exist_ok=True)
                                        else:
                                            logging.debug(f"              Skipping empty subsection name under Section: {section}")
                                elif subsection_list is not None: # Log if it's not a list but not None either
                                     logging.warning(f"            Expected list of subsections for {header}/{subheader}/{section}, found: {type(subsection_list)}. Value: {subsection_list}")

                        elif section_dict is not None: # Log if not a dict but not None
                             logging.warning(f"        Expected dict for sections under {header}/{subheader}, found: {type(section_dict)}. Value: {section_dict}")

                elif isinstance(sub_level_data, list): # Flat: Header -> Subheader list
                    logging.debug(f"    Processing Flat Header: {header}")
                    for subheader_item in sub_level_data:
                        if subheader_item: # Skip if item name is empty string or None
                             subheader_path = os.path.join(header_path, subheader_item)
                             logging.debug(f"      Checking/Creating Flat Subheader Item: {subheader_path}")
                             os.makedirs(subheader_path, exist_ok=True)
                        else:
                             logging.debug(f"      Skipping empty flat subheader item under Header: {header}")
                else:
                     logging.warning(f"  Unknown structure type for header '{header}': {type(sub_level_data)}. Value: {sub_level_data}")

            logging.info(f"Structure verification complete for: {safe_company_name}")

        except Exception as e:
            logging.error(f"Failed to create company structure for '{company_name}': {e}", exc_info=True)
            # Re-raise the exception so the calling function knows structure creation failed
            raise

    # Inside FileArchiveApp class

    # Inside FileArchiveApp class

    def perform_file_upload(self, company_name, header, subheader, section, subsection,
                            source_file_path, # Renamed for clarity
                            intended_destination_filename # New argument
                            ):
        """
        Performs the actual file upload logic using a pre-determined destination filename.
        Handles path creation including subsections and backups.

        Args:
            company_name (str): The display name of the company.
            header (str): Selected header.
            subheader (str): Selected subheader.
            section (str): Selected section.
            subsection (str): Selected subsection.
            source_file_path (str): The full path to the source file to upload.
            intended_destination_filename (str): The final filename to use in the archive.

        Returns:
            True on success.
            Raises Exception on errors caught during IO.
        """
        dest_path = "Unknown" # Initialize for logging
        safe_company_name = "Unknown" # Initialize

        # Log entry point and arguments clearly
        logging.info(f"[UploadLogicV2 ENTRY] Args: Co='{company_name}', H='{header}', S='{subheader}', Sec='{section}', SubSec='{subsection}', SrcPath='{source_file_path}', DestName='{intended_destination_filename}'")

        try:
            # --- Ensure company info is set ---
            if not hasattr(self, 'current_company') or self.current_company.get("display_name") != company_name:
                 self.create_company_structure(company_name)
            safe_company_name = self.current_company["safe_name"]
            logging.debug(f"[UploadLogicV2] Using safe company name: {safe_company_name}")

            # --- Determine structure and path components ---
            structure_options = self.structure.get(header, [])
            has_subsections_defined = False
            if isinstance(structure_options, dict): # Nested Header
                section_dict = structure_options.get(subheader, {})
                if section and section in section_dict:
                    subsections_list = section_dict.get(section, [])
                    has_subsections_defined = bool(subsections_list)
            # (Prefix calculation removed - no longer needed here)

            # --- Calculate Destination Path ---
            dest_path = os.path.join(self.archives_path, safe_company_name, header)
            if isinstance(structure_options, dict): # Nested Header path construction
                if subheader: dest_path = os.path.join(dest_path, subheader)
                if section:
                    dest_path = os.path.join(dest_path, section)
                    if subsection and has_subsections_defined:
                        dest_path = os.path.join(dest_path, subsection)
            elif subheader: # Flat structure path construction
                dest_path = os.path.join(dest_path, subheader)
            logging.debug(f"[UploadLogicV2] Calculated destination folder: {dest_path}")

            # --- Ensure Destination Directory Exists ---
            try:
                os.makedirs(dest_path, exist_ok=True)
                logging.debug(f"[UploadLogicV2] Ensured destination directory exists: {dest_path}")
            except Exception as e_mkdir:
                logging.error(f"[UploadLogicV2] FAILED to create destination directory {dest_path}: {e_mkdir}", exc_info=True)
                raise IOError(f"Failed to create directory: {dest_path}") from e_mkdir

            # --- Final Destination File Path ---
            # Use the filename passed into the function
            dest_file = os.path.join(dest_path, intended_destination_filename)
            logging.debug(f"[UploadLogicV2] Final destination file path: {dest_file}")

            # --- Backup Logic (uses intended_destination_filename) ---
            if os.path.exists(dest_file):
                logging.warning(f"[UploadLogicV2] Destination file exists: {dest_file}. Creating backup.")
                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                base, ext = os.path.splitext(intended_destination_filename)
                backup_name = f"{base}_backup_{timestamp}{ext}"
                backup_path = os.path.join(dest_path, backup_name)
                try:
                    count = 0
                    while os.path.exists(backup_path): # Handle collision
                         count += 1
                         backup_name = f"{base}_backup_{timestamp}_{count}{ext}"
                         backup_path = os.path.join(dest_path, backup_name)
                    os.rename(dest_file, backup_path)
                    logging.info(f"[UploadLogicV2] Existing file versioned: {dest_file} -> {backup_path}")
                except (PermissionError, FileNotFoundError, OSError) as e_mv:
                    logging.error(f"[UploadLogicV2] FAILED to version existing file {dest_file}: {e_mv}", exc_info=True)
                    raise IOError(f"Error versioning existing file '{intended_destination_filename}'") from e_mv

            # --- Copy File ---
            try:
                logging.debug(f"[UploadLogicV2] Attempting copy: '{source_file_path}' -> '{dest_file}'")
                shutil.copy2(source_file_path, dest_file) # Use copy2
                logging.info(f"[UploadLogicV2] File copied successfully: {source_file_path} -> {dest_file}")
                return True # Indicate success
            except Exception as e_copy:
                 logging.error(f"[UploadLogicV2] FAILED to copy file '{source_file_path}' to '{dest_file}': {e_copy}", exc_info=True)
                 raise IOError(f"Failed to copy file to destination") from e_copy

        except Exception as e_main:
            # Catch any other unexpected errors
            logging.error(f"[UploadLogicV2 FATAL] Error processing {source_file_path} -> {dest_path}/{intended_destination_filename}: {e_main}", exc_info=True)
            raise e_main # Re-raise exception

    # --------------------------------------------------------------------------
    # Dashboard and Search Functionality
    # --------------------------------------------------------------------------
    def open_dashboard(self):
        dashboard = ctk.CTkToplevel(self.main_app)
        dashboard.transient(self.main_app) # Stay on top
        dashboard.title("Dashboard")
        self.center_window(dashboard, 500, 400)
        dashboard.grab_set()

        total_files, total_size, file_types = 0, 0, {}
        for root, _, files in os.walk(self.archives_path):
            for f in files:
                total_files += 1
                fpath = os.path.join(root, f)
                try:
                    size = os.path.getsize(fpath)
                    total_size += size
                except Exception:
                    continue
                ext = os.path.splitext(f)[1].lower()
                file_types[ext] = file_types.get(ext, 0) + 1
        stats_text = f"Total Files: {total_files}\nTotal Size: {total_size} bytes\nFile Types:\n"
        for ext, count in file_types.items():
            stats_text += f"  {ext or 'no ext'}: {count}\n"
        ctk.CTkLabel(dashboard, text=get_translation("ctklabel_text_archive_statistics"), font=("Segoe UI", 16, "bold")).pack(pady=5)
        ctk.CTkLabel(dashboard, text=stats_text, font=("Segoe UI", 12)).pack(pady=5)
        ctk.CTkLabel(dashboard, text=get_translation("ctklabel_text_search_analytics"), font=("Segoe UI", 16, "bold")).pack(pady=5)
        with self.search_queries_lock:
            analytics_text = "\n".join(self.search_queries) if self.search_queries else "No searches performed yet."
        ctk.CTkLabel(dashboard, text=analytics_text, font=("Segoe UI", 12)).pack(pady=5)



    # --------------------------------------------------------------------------
    # Application Closing
    # --------------------------------------------------------------------------
    def on_closing(self):
        """Handle application closing efficiently without blocking"""
        try:
            # Log the user logout
            if self.current_user:
                logging.info(f"User '{self.current_user['username']}' logged out")

            # Stop watchdog observer in a non-blocking way
            if hasattr(self, 'observer') and self.observer.is_alive():
                self.observer.stop()
                # Don't join the thread - it can cause hanging

            # Hide archive folder without recursion (which can be slow)
            if platform.system() == "Windows":
                try:
                    # Use a non-blocking approach - start process and don't wait
                    subprocess.Popen(['attrib', '+h', '+s', self.archives_path],
                                shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
                except Exception as e:
                    logging.error(f"Error hiding archive folder: {e}")
            else:
                try:
                    # Use a non-blocking approach
                    subprocess.Popen(['chmod', '700', self.archives_path])
                except Exception as e:
                    logging.error(f"Error changing archive folder permissions: {e}")

            logging.info("Application closing")
        except Exception as e:
            logging.error(f"Error during application cleanup: {e}")
        finally:
            # Always destroy the window, even if cleanup fails
            if hasattr(self, 'dnd_root'):
                self.dnd_root.destroy()
            else:
                self.main_app.destroy()

    def cleanup_admin_resources(self):
        """Clean up any resources created during admin session"""
        import os
        import shutil

        # Define paths that need cleanup
        temp_paths = []

        # Add any admin-specific temporary directories or files
        if hasattr(self, 'temp_archive_path') and self.temp_archive_path:
            temp_paths.append(self.temp_archive_path)

        # You can add more paths as needed

        # Clean up each path
        for path in temp_paths:
            try:
                if os.path.isdir(path):
                    shutil.rmtree(path)
                    logging.info(f"Removed temporary directory: {path}")
                elif os.path.isfile(path):
                    os.remove(path)
                    logging.info(f"Removed temporary file: {path}")
            except Exception as e:
                logging.error(f"Error removing {path}: {e}")

    # --------------------------------------------------------------------------
    # Run the Application
    # --------------------------------------------------------------------------
    def run(self):
        """Run the application with an enhanced splash screen"""
        # Create a stylish splash screen
        if hasattr(self, 'dnd_root'):
            splash = ctk.CTkToplevel(self.dnd_root)
        else:
            splash = ctk.CTkToplevel(self.main_app)
        splash.transient(self.main_app) # Stay on top

        splash.title("")
        splash.attributes("-topmost", True)
        splash.overrideredirect(True)  # Remove window decorations
        self.center_window(splash, 500, 300)

        # Add content to splash screen with better styling
        splash_frame = ctk.CTkFrame(splash, corner_radius=15, border_width=2,
                                border_color=["#565B5E", "#565B5E"])
        splash_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Logo placeholder (in real app, use an actual logo image)
        logo_label = ctk.CTkLabel(splash_frame, text=get_translation("ctklabel_text_empty_string"), font=("Segoe UI", 48))
        logo_label.pack(pady=(40, 5))

        # App title with larger font
        ctk.CTkLabel(splash_frame, text=get_translation("ctklabel_text_file_archiving_system"),
                    font=("Segoe UI", 28, "bold")).pack(pady=(0, 5))

        # Version info
        ctk.CTkLabel(splash_frame, text=get_translation("ctklabel_text_version_10"),
                    font=("Segoe UI", 14)).pack(pady=(0, 20))

        # Loading message
        loading_label = ctk.CTkLabel(splash_frame, text=get_translation("ctklabel_text_initializing"),
                                    font=("Segoe UI", 14))
        loading_label.pack(pady=(0, 10))

        # Stylish progress bar
        progress = ctk.CTkProgressBar(splash_frame, width=400, height=15,
                                    corner_radius=5)
        progress.pack(pady=10)
        progress.set(0)

        # Simulate loading progress with changing messages
        loading_messages = [
            "Loading interface...",
            "Checking security...",
            "Initializing storage...",
            "Preparing document structure...",
            "Starting watchdog monitors...",
            "Ready to launch!"
        ]

        def update_splash(i=0):
            if i < len(loading_messages):
                progress.set((i+1)/len(loading_messages))
                loading_label.configure(text=loading_messages[i])
                splash.after(400, lambda: update_splash(i+1))
            else:
                # Close splash and show login screen
                splash.destroy()
                # NOW show the login dialog AFTER splash screen
                self.authenticate_user()

        # Start the splash screen updates
        splash.after(500, update_splash)

        # Start the mainloop
        if hasattr(self, 'dnd_root'):
            self.dnd_root.mainloop()
        else:
            self.main_app.mainloop()

if __name__ == '__main__':
    app = FileArchiveApp()
    app.run() # Run your application