# Comprehensive Explanation of the File Archiving Application

## 1. Introduction

This Python application is a graphical user interface (GUI) tool designed for managing and organizing files within a hierarchical archive system. It provides features for user authentication, role-based access control (administrators and regular users), multi-language support (English and Arabic), and various file operations including uploads, backups, previews, and version rollbacks.

## 2. Architecture Overview

The application follows an MVC-like (Model-View-Controller) pattern:

*   **`FileArchiveApp` (in `test.py`)**: This is the main class, acting as the primary View and part of the Controller. It initializes and manages the `customtkinter`-based GUI, handles user interactions and sessions, maintains application state, and orchestrates operations with other controllers.
*   **Controllers (`controllers/` directory):**
    *   `ArchiveController.py`: Manages the logic related to the archive's folder structure, including dynamic discovery of folders.
    *   `UserController.py`: Handles user authentication, password changes, and interfaces with the user data store.
*   **Data and Supporting Files:**
    *   `users.db`: A SQLite database for persistent storage of user credentials (usernames, hashed passwords, roles).
    *   `translations.json`: Stores UI text strings for internationalization (English and Arabic).
    *   `translations.py`: Contains the logic to load and manage these translations.
    *   `archive_app.log`: Records application events for debugging and activity tracking.

## 3. User Interface & Experience

The application features a modern, tabbed interface:

*   **Login:** On startup, a splash screen appears, followed by a modal login dialog. Credentials are verified by `UserController` against `users.db`.
*   **Main Window:**
    *   **Header:** Displays the application title and current user information.
    *   **Tabs:**
        *   **Upload Tab:** Allows users to select a Company, Header, Subheader, Section, and Subsection to define the archive path. It supports single file uploads (with interactive naming convention checks), batch uploads, a "Scan & Archive" feature (using WIA on Windows), and a drag-and-drop area. Admins can also create new folders (Sections/Subsections) from this tab.
        *   **Manage Tab:** Enables users to preview archived files (images directly, others via the OS default application; printing is supported for previews) and rollback files to previous backup versions. Admins see a "Recent Activity Log" here.
        *   **Settings Tab:** Users can switch UI themes (Dark, Light, System; also via Ctrl+T). Admins can change their own passwords here.
        *   **Admin Tab (Admin Only):** Provides administrative functions:
            *   *System Management:* Refresh folder visibility, search the archive, open a statistics dashboard.
            *   *User Management:* Add, edit (roles, passwords), and delete users.
            *   *Statistics:* View system and user statistics.
    *   **Status Bar:** Displays notifications, progress of background tasks, and error messages.
*   **Internationalization (i18n):** Supports English and Arabic. The UI dynamically updates when the language is switched.
*   **Theming:** Offers Dark, Light, and System theme options.
*   **Responsiveness:** Long-running tasks (e.g., search, batch uploads, admin folder creation) are executed in background threads using a `ThreadPoolExecutor`, with UI updates (progress bars, status messages) handled safely via a `ui_queue` to prevent the application from freezing.
*   **Interactive Dialogs:** Uses `customtkinter` and standard `tkinter` dialogs for login, file operations, confirmations, and user input.

## 4. Core Functionality - File Archiving

*   **Archive Hierarchy (`self.structure`):** A dictionary in `FileArchiveApp` defines a template for the archive's hierarchical structure (Header, Subheader, Section, Subsection).
*   **Dynamic Folder Discovery (`ArchiveController.get_dynamic_folder_options`):** This method populates UI dropdowns by scanning for existing folders on disk and merging them with the `self.structure` template, allowing flexibility.
*   **Upload Process (`FileArchiveApp.perform_file_upload` initiated by `upload_file`):**
    *   The destination path is determined by user selections (Company, Header, etc.).
    *   A `required_prefix` (e.g., the name of the subsection) is determined for filenames based on the selected archive depth.
    *   **Naming Convention (Single Upload):** If a file doesn't meet the prefix rule, the user is prompted to auto-rename it (prefix added), manually rename it, or cancel the upload.
    *   **Backup Creation:** If a file with the same name exists at the destination, the existing file is renamed with a timestamp (e.g., `Filename_backup_YYYYMMDDHHMMSS.ext`) before the new file is saved.
*   **Batch Upload (`FileArchiveApp.batch_upload`):**
    *   Handles multiple file uploads.
    *   Pre-checks naming conventions. Users are prompted to auto-rename non-compliant files, skip them, or cancel the batch. Uploads are threaded.
*   **Drag & Drop (`FileArchiveApp.on_drop`):**
    *   Processes files dropped onto the Upload tab.
    *   Strictly rejects files that don't match the `required_prefix` for the current UI selection (no renaming offered). Uploads are threaded.
*   **Scan & Archive (`FileArchiveApp.scan_and_archive`):**
    *   Uses WIA (Windows Image Acquisition) for scanning.
    *   Prompts the user for a filename suffix; the system automatically adds the `required_prefix`.
    *   Saves with backup logic.
*   **Admin Folder Creation (`FileArchiveApp.add_structure_element_dialog_contextual`):**
    *   Admins can add new Sections or Subsections to the live archive structure.
    *   The new folder is created on disk in a background task, and UI dropdowns refresh to include it.

## 5. Core Functionality - User Management

*   **`UserController`**: Manages user authentication and password changes.
*   **Authentication (`UserController.verify_credentials`):**
    *   Checks username and verifies the provided password against a securely stored hash using `passlib.hash.pbkdf2_sha256.verify()`. Plain-text passwords are never stored.
*   **Data Persistence (`users.db`):**
    *   User details (username, hashed password, role) are stored in the `users.db` SQLite database.
    *   `FileArchiveApp` handles initializing the DB, loading users into an in-memory dictionary at startup, and saving changes back to the DB.
    *   A default admin user is created if none exist.
*   **Password Management (`UserController.change_password`):**
    *   Admins can change their own passwords via the Settings tab. The process involves verifying the current password and then updating the stored hash.
*   **Admin User Operations (Admin Tab in `FileArchiveApp`):**
    *   **Add User:** Admins can create new users, assigning roles and initial passwords (which are then hashed).
    *   **Edit User:** Admins can change existing users' roles or reset their passwords.
    *   **Delete User:** Admins can remove user accounts (but not their own).
    *   All changes are reflected in both the in-memory user store and `users.db`.

## 6. Key Supporting Modules

*   **`translations.py` / `translations.json`**: Provide internationalization. `.json` stores translations; `.py` loads and manages them, including fallbacks.
*   **`watchdog` library**: If available, monitors the archive folder for real-time changes, updating a UI notification label via the `ui_queue`.
*   **`ThreadPoolExecutor`**: Manages a thread pool for background tasks (batch uploads, search, etc.), using a `ui_queue` for safe GUI updates from these threads, ensuring UI responsiveness.
*   **`logging` module**: Configured to write application events (INFO level and above) to `archive_app.log`, crucial for debugging and activity tracking. The Admin Tab's "Recent Activity Log" displays recent entries from this file.

## 7. Project Structure & Future Directions

*   **Structure:** The project is organized with `test.py` as the main application entry point, a `controllers` directory for business logic, and separate files for data (`.json`, `.db`) and specific functionalities like translations.
*   **Potential Refinements:**
    *   **Configuration Management:** Move hardcoded settings (like `self.structure`, default passwords) to external configuration files.
    *   **User Store Consistency:** Further centralize user database interactions to ensure the in-memory `users` dictionary and `users.db` are always synchronized.
    *   **Code Refactoring (DRY):** Consolidate repeated logic (e.g., `required_prefix` determination) into shared utilities.
    *   **Enhanced Security:** Conduct a comprehensive security review, especially for file system permissions and access controls in a production setting.
    *   **Error Handling:** Make error messages more specific and user-friendly, particularly for I/O and background task issues.
    *   **Test Coverage:** Implement dedicated unit and integration tests to improve reliability and facilitate safer code changes.

## 8. Conclusion

This application provides a comprehensive solution for structured file archiving with a user-friendly graphical interface. It effectively separates concerns, handles user management securely, and offers a good degree of flexibility through dynamic folder discovery and administrative controls. The use of background threading and internationalization further enhances the user experience. The codebase forms a solid foundation that can be extended and refined for even greater robustness and functionality.
