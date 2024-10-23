import winreg

def remove_context_menu():
    try:
        # Remove context menu for all files
        key_path = r'Software\Classes\*\shell\Study Aggregator'
        remove_registry_entry(key_path)

        # Remove context menu for folders
        key_path = r'Software\Classes\Directory\shell\Study Aggregator'
        remove_registry_entry(key_path)

        # Remove context menu for drives (e.g., CDs)
        key_path = r'Software\Classes\Drive\shell\Study Aggregator'
        remove_registry_entry(key_path)

        # Remove context menu for .zip files
        key_path = r'Software\Classes\SystemFileAssociations\.zip\shell\Study Aggregator'
        remove_registry_entry(key_path)

        print("Context menu options removed successfully.")
    except Exception as e:
        print(f"Failed to modify the registry: {e}")

def remove_registry_entry(key_path):
    """Helper function to remove a context menu entry"""
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path + r'\command')
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, key_path)
        print(f"Removed context menu entry at {key_path}")
    except FileNotFoundError:
        print(f"Registry entry not found: {key_path}")
    except Exception as e:
        print(f"Error removing registry entry: {e}")

if __name__ == "__main__":
    remove_context_menu()
