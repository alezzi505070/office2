import os
import platform
import time

print("--- Starting File Write Test ---")

# Determine the exact same log directory as your app
if platform.system() == "Windows":
    log_dir = os.path.join(os.getenv('APPDATA'), 'FileArchiveApp')
else:
    log_dir = os.path.join(os.path.expanduser('~'), '.FileArchiveApp')

print(f"Target directory: {log_dir}")

# Ensure the directory exists
try:
    os.makedirs(log_dir, exist_ok=True)
    print("Directory exists or was created successfully.")
except Exception as e:
    print(f"CRITICAL ERROR creating directory: {e}")
    # If this fails, the problem is with creating the directory itself.
    # This could be a permissions issue.
    exit()

# Define the file path
test_file_path = os.path.join(log_dir, 'test_write.txt')
print(f"Attempting to write to: {test_file_path}")

# Try to write to the file
try:
    with open(test_file_path, 'w', encoding='utf-8') as f:
        f.write(f"This is a test write at {time.ctime()}.\n")
        f.write("If you can read this, file writing is possible.\n")
    
    print("\nSUCCESS: File was written successfully!")
    print(f"Please check the contents of '{test_file_path}'")

except Exception as e:
    print(f"\nFAILURE: Could not write to the file.")
    print(f"Error details: {e}")
    print("\nThis indicates a potential permissions problem, an issue with your antivirus software, or a problem with the file path itself.")

print("--- Test Complete ---")