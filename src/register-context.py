import os
import sys
import winreg
from pathlib import Path


def register_context_menu():
    """
    Register context menu items for both files and folders.
    Files: "Merge PDF (files)"
    Folders: "Merge PDF (folder)"
    """
    python_exe = sys.executable
    script_path = os.path.abspath("merge-pdf.py")
    icon_path = os.path.abspath("icon.ico")
    
    # Check if custom icon exists, otherwise use Python icon
    if not os.path.exists(icon_path):
        icon_path = python_exe
        print(f"Warning: icon.ico not found, using Python icon instead")
    
    # Registry keys (system-wide)
    file_key = r"*\shell\Merge PDF (files)"
    file_cmd_key = file_key + r"\command"

    dir_key = r"Directory\shell\Merge PDF (folder)"
    dir_cmd_key = dir_key + r"\command"

    try:
        # File right-click option (multi-selection enabled)
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, file_key) as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Merge PDF (files)")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon_path)
            winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Document")

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, file_cmd_key) as cmd:
            command = f'"{python_exe}" "{script_path}" -auto "%V"'
            winreg.SetValueEx(cmd, None, 0, winreg.REG_SZ, command)

        # Folder right-click option
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, dir_key) as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Merge PDF (folder)")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, icon_path)

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, dir_cmd_key) as cmd:
            command = f'"{python_exe}" "{script_path}" -auto "%V"'
            winreg.SetValueEx(cmd, None, 0, winreg.REG_SZ, command)

        print("Context menu registered successfully (system-wide).")
        print("   - File right-click: Merge PDF (files)")
        print("   - Folder right-click: Merge PDF (folder)")
        print(f"   - Icon: {Path(icon_path).name}")
        print("\nNOTE: Windows only supports .ico format for context menu icons.")
        print("   For best results, convert icon.png to icon.ico")
        print("\nPlease restart File Explorer to see changes:")
        print("   1. Press Ctrl+Shift+Esc (Task Manager)")
        print("   2. Find 'Windows Explorer' -> Right-click -> Restart")
        input("\nPress Enter to exit...")

    except PermissionError:
        print("ERROR: Permission denied. Please run this script as Administrator.")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"ERROR: Failed to register context menu: {e}")
        input("Press Enter to exit...")


def unregister_context_menu():
    """
    Remove all previously registered context menu entries.
    """
    try:
        deleted_count = 0
        for key_path in [
            r"*\shell\Merge PDF (files)",
            r"Directory\shell\Merge PDF (folder)"
        ]:
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path + r"\command")
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
                deleted_count += 1
            except FileNotFoundError:
                pass
        
        if deleted_count > 0:
            print(f"Context menu unregistered successfully ({deleted_count} entries removed).")
        else:
            print("INFO: No context menu entries found to remove.")
        
        input("Press Enter to exit...")
    except PermissionError:
        print("ERROR: Permission denied. Please run this script as Administrator.")
        input("Press Enter to exit...")
    except Exception as e:
        print(f"ERROR: Failed to unregister context menu: {e}")
        input("Press Enter to exit...")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Register or unregister context menu for merge-pdf."
    )
    parser.add_argument("--uninstall", action="store_true",
                        help="Unregister the context menu.")
    args = parser.parse_args()

    if args.uninstall:
        unregister_context_menu()
    else:
        register_context_menu()