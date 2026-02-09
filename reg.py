import os
import subprocess
import winreg

AUMID = "Ronsin.StudyAggregator"


def add_context_menu():
    # Path to your EXE in the %appdata%\Study Aggregator folder
    exe_path = os.path.join(os.getenv('APPDATA'), 'Study Aggregator', 'Study Aggregator.exe')

    # Context menu entry name and command to run
    menu_name = "Study Aggregator"
    command = f'"{exe_path}" "%1"'

    try:
        # Add context menu for all files
        key_path = r'Software\Classes\*\shell\Study Aggregator'
        add_registry_entry(key_path, menu_name, command, exe_path)

        # Add context menu for folders
        key_path = r'Software\Classes\Directory\shell\Study Aggregator'
        add_registry_entry(key_path, menu_name, command, exe_path)

        # Add context menu for drives (e.g., CDs)
        key_path = r'Software\Classes\Drive\shell\Study Aggregator'
        add_registry_entry(key_path, menu_name, command, exe_path)

        # Add context menu for .zip files
        key_path = r'Software\Classes\SystemFileAssociations\.zip\shell\Study Aggregator'
        add_registry_entry(key_path, menu_name, command, exe_path)

        print(f"Context menu options added successfully with EXE: {exe_path} as the icon source.")
    except Exception as e:
        print(f"Failed to modify the registry: {e}")


def add_registry_entry(key_path, menu_name, command, exe_path):
    """Helper function to create a context menu entry with icon from the EXE"""
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, '', 0, winreg.REG_SZ, menu_name)
            # Use the EXE file to retrieve the icon, using index 0
            winreg.SetValueEx(key, 'Icon', 0, winreg.REG_SZ, f'{exe_path},0')

        # Create the command subkey
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path + r'\command') as command_key:
            winreg.SetValueEx(command_key, '', 0, winreg.REG_SZ, command)

        print(f"Added context menu entry at {key_path} with EXE icon: {exe_path}")
    except Exception as e:
        print(f"Error adding registry entry at {key_path}: {e}")


def add_scheduled_update_check():
    """Create a scheduled task to check for updates every 10 minutes when a user is logged on."""
    checker_path = os.path.join(os.getenv('APPDATA'), 'Study Aggregator', 'update_checker.exe')
    task_name = "StudyAggregatorUpdateCheck"

    try:
        result = subprocess.run(
            [
                "schtasks", "/create",
                "/tn", task_name,
                "/tr", f'"{checker_path}"',
                "/sc", "minute",
                "/mo", "10",
                "/rl", "LIMITED",
                "/it",
                "/f",
            ],
            capture_output=True,
            text=True,
        )
        if result.returncode == 0:
            print(f"Scheduled task '{task_name}' created successfully (every 10 minutes).")
        else:
            print(f"Failed to create scheduled task: {result.stderr.strip()}")
    except Exception as e:
        print(f"Error creating scheduled task: {e}")


def add_aumid():
    """Register the AUMID so Windows toast notifications work with interactive buttons."""
    icon_path = os.path.join(os.getenv('APPDATA'), 'Study Aggregator', 'agg.ico')
    key_path = rf"Software\Classes\AppUserModelId\{AUMID}"

    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "DisplayName", 0, winreg.REG_SZ, "Study Aggregator")
            winreg.SetValueEx(key, "IconUri", 0, winreg.REG_SZ, icon_path)
        print(f"AUMID '{AUMID}' registered successfully.")
    except Exception as e:
        print(f"Error registering AUMID: {e}")


if __name__ == "__main__":
    add_context_menu()
    add_scheduled_update_check()
    add_aumid()
