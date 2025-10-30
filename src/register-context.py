import os
import sys
import winreg


def register_context_menu():
    """
    Register context menu items for both files and folders.
    Files: "Merge PDF (files)"
    Folders: "Merge PDF (folder)"
    """
    python_exe = sys.executable  # Path to current Python executable
    script_path = os.path.abspath("merge-pdf.py")

    # Registry keys (system-wide)
    file_key = r"*\shell\Merge PDF (files)"
    file_cmd_key = file_key + r"\command"

    dir_key = r"Directory\shell\Merge PDF (folder)"
    dir_cmd_key = dir_key + r"\command"

    try:
        # File right-click option (multi-selection enabled)
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, file_key) as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Merge PDF (files)")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, python_exe)
            winreg.SetValueEx(key, "MultiSelectModel", 0, winreg.REG_SZ, "Document")

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, file_cmd_key) as cmd:
            # Use %V for proper path handling with spaces
            command = f'"{python_exe}" "{script_path}" -auto "%V"'
            winreg.SetValueEx(cmd, None, 0, winreg.REG_SZ, command)

        # Folder right-click option
        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, dir_key) as key:
            winreg.SetValueEx(key, None, 0, winreg.REG_SZ, "Merge PDF (folder)")
            winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, python_exe)

        with winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, dir_cmd_key) as cmd:
            # Use %V for proper path handling
            command = f'"{python_exe}" "{script_path}" -auto "%V"'
            winreg.SetValueEx(cmd, None, 0, winreg.REG_SZ, command)


        print("Context menu registered successfully (system-wide).")
        print(" - File right-click: Merge PDF (files)")
        print(" - Folder right-click: Merge PDF (folder)")
        print("\nNOTE: For multiple file selection, select files one by one while holding Ctrl.")
        print("Please restart File Explorer if you do not see the menu.")
        input("\nPress Enter to exit...")

    except PermissionError:
        print("Permission denied. Please run this script as Administrator.")
        input("Press Enter to exit...")
    except Exception as e:
        print("Failed to register context menu:", e)
        input("Press Enter to exit...")


def unregister_context_menu():
    """
    Remove all previously registered context menu entries.
    """
    try:
        for key_path in [
            r"*\shell\Merge PDF (files)",
            r"Directory\shell\Merge PDF (folder)"
        ]:
            try:
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path + r"\command")
                winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, key_path)
            except FileNotFoundError:
                pass
        print("Context menu unregistered successfully.")
        input("Press Enter to exit...")
    except PermissionError:
        print("Permission denied. Please run this script as Administrator.")
        input("Press Enter to exit...")
    except Exception as e:
        print("Failed to unregister context menu:", e)
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