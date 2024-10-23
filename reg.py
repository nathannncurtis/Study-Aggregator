import os
import winreg

def add_context_menu():
    # Path to your EXE in the AppData\Roaming\Study Aggregator folder
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

if __name__ == "__main__":
    add_context_menu()
