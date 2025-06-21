import logging
from passlib.hash import pbkdf2_sha256
from tkinter import messagebox

class UserController:
    """
    Handles user authentication and password management.
    """
    def __init__(self, users_store):
        # users_store: dict of username -> {password: hash, role: str}
        self._users = users_store

    def verify_credentials(self, username, password):
        """
        Verifies username/password against the users store.
        Returns a dict {'username': ..., 'role': ...} on success, or None on failure.
        """
        if username in self._users and pbkdf2_sha256.verify(password, self._users[username]["password"]):
            return {"username": username, "role": self._users[username]["role"]}
        return None

    def change_password(self, current_user, current_pwd, new_pwd, confirm_pwd, parent_window=None):
        """
        Changes the password for the current_user (admin only).
        Returns True on success, False otherwise.
        """
        if not current_user:
            messagebox.showerror("Error", "No user is logged in.", parent=parent_window)
            return False

        # Only admins may change their own password
        if current_user.get("role", "").lower() != "admin":
            messagebox.showerror("Permission Denied", "Only admin users can change their password.", parent=parent_window)
            return False

        if not (current_pwd and new_pwd and confirm_pwd):
            messagebox.showerror("Error", "All fields are required", parent=parent_window)
            return False

        if new_pwd != confirm_pwd:
            messagebox.showerror("Error", "New passwords do not match", parent=parent_window)
            return False

        username = current_user["username"]
        if not pbkdf2_sha256.verify(current_pwd, self._users[username]["password"]):
            messagebox.showerror("Error", "Current password is incorrect", parent=parent_window)
            return False

        # Update hash
        self._users[username]["password"] = pbkdf2_sha256.hash(new_pwd)
        messagebox.showinfo("Success", "Password changed successfully", parent=parent_window)
        logging.info(f"Admin '{username}' changed their password")
        return True