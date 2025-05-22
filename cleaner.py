import os
import shutil
import tempfile
import ctypes
import winshell
from win32com.client import Dispatch

def run_as_admin():
    """Check if the script is running as admin, if not relaunch with admin privileges"""
    if ctypes.windll.shell32.IsUserAnAdmin():
        return True
    
    script = os.path.abspath(__file__)
    params = ' '.join([script] + sys.argv[1:])
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, params, None, 1)
    return False

def clean_temp_folders():
    """Clean various temporary folders in Windows"""
    # System temp folder
    temp_folder = tempfile.gettempdir()
    print(f"Cleaning system temp folder: {temp_folder}")
    for filename in os.listdir(temp_folder):
        file_path = os.path.join(temp_folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)
        except Exception as e:
            print(f"Failed to delete {file_path}. Reason: {e}")

    # Windows temp folder
    win_temp = r'C:/Windows/Temp'
    print(f"Cleaning Windows temp folder: {win_temp}")
    if os.path.exists(win_temp):
        for filename in os.listdir(win_temp):
            file_path = os.path.join(win_temp, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

    # User temp folders
    user_temp = os.environ.get('USERPROFILE') + r'/AppData/Local/Temp'
    print(f"Cleaning user temp folder: {user_temp}")
    if os.path.exists(user_temp):
        for filename in os.listdir(user_temp):
            file_path = os.path.join(user_temp, filename)
            try:
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

def clean_recycle_bin():
    """Empty the recycle bin"""
    print("Emptying recycle bin...")
    try:
        winshell.recycle_bin().empty(confirm=False, show_progress=False, sound=False)
        print("Recycle bin emptied successfully.")
    except Exception as e:
        print(f"Failed to empty recycle bin. Reason: {e}")

def clean_prefetch():
    """Clean prefetch folder (requires admin)"""
    prefetch_path = r'C:/Windows/Prefetch'
    print(f"Cleaning prefetch folder: {prefetch_path}")
    if os.path.exists(prefetch_path):
        for filename in os.listdir(prefetch_path):
            file_path = os.path.join(prefetch_path, filename)
            try:
                if os.path.isfile(file_path):
                    os.unlink(file_path)
            except Exception as e:
                print(f"Failed to delete {file_path}. Reason: {e}")

def main():
    # First check for admin rights
    if not run_as_admin():
        print("This script requires administrator privileges to perform all cleaning tasks.")
        print("The script will now exit without performing any cleaning operations.")
        input("Press Enter to exit...")
        sys.exit(1)
    
    # Only proceed if we have admin rights
    print("Starting Windows 10 optimization cleanup...")
    
    # Perform cleaning tasks
    clean_temp_folders()
    clean_recycle_bin()
    clean_prefetch()
    
    print("Cleanup completed successfully!")
    print("It's recommended to restart your computer for optimal performance.")

if __name__ == "__main__":
    import sys
    main()