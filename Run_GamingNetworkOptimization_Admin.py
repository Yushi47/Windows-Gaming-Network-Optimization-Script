import ctypes
import sys
import os
import subprocess # Added for running PowerShell to get adapters
import json       # Added for parsing PowerShell output

# --- Configuration ---
TARGET_PS_SCRIPT_NAME = "GamingNetworkOptimization.ps1"

def get_network_adapters_from_powershell():
    """
    Uses PowerShell to get a list of network adapters with their details.
    This part does NOT require admin rights.
    """
    print("Fetching available network adapters...")
    # Command to get adapter info as JSON, filtering out those without a GUID
    command = "Get-NetAdapter | Select-Object Name, InterfaceDescription, InterfaceGuid | Where-Object {$_.InterfaceGuid -ne $null} | ConvertTo-Json -Compress"
    try:
        # Run PowerShell command, hide its window for this data retrieval step
        process = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", command],
            capture_output=True, text=True, check=True, encoding='utf-8',
            creationflags=subprocess.CREATE_NO_WINDOW 
        )
        adapters_json = process.stdout
        adapters_data = json.loads(adapters_json)
        # If PowerShell returns a single adapter, ConvertTo-Json might output a single object instead of an array
        if isinstance(adapters_data, dict):
            return [adapters_data] # Ensure it's always a list
        return adapters_data
    except subprocess.CalledProcessError as e:
        error_msg = f"Error getting network adapters from PowerShell: {e}\nStderr: {e.stderr}"
        print(error_msg)
        ctypes.windll.user32.MessageBoxW(None, error_msg, "Adapter Enumeration Error", 0x10 | 0x0)
        return None
    except json.JSONDecodeError as e:
        error_msg = f"Error decoding JSON from PowerShell adapter list: {e}\nRaw output: {process.stdout if 'process' in locals() else 'N/A'}"
        print(error_msg)
        ctypes.windll.user32.MessageBoxW(None, error_msg, "Adapter Enumeration Error", 0x10 | 0x0)
        return None
    except Exception as e:
        error_msg = f"An unexpected error occurred while getting adapters: {e}"
        print(error_msg)
        ctypes.windll.user32.MessageBoxW(None, error_msg, "Adapter Enumeration Error", 0x10 | 0x0)
        return None

def select_guids_for_tweaks(adapters):
    """
    Prompts the user to select one or more adapters from the provided list.
    Returns a list of selected GUID strings.
    """
    if not adapters:
        print("No network adapters with GUIDs were found or an error occurred.")
        return []

    print("\nAvailable Network Adapters (select for interface-specific tweaks):")
    for i, adapter in enumerate(adapters):
        name = adapter.get('Name', 'N/A')
        desc = adapter.get('InterfaceDescription', 'N/A')
        guid = adapter.get('InterfaceGuid', 'N/A')
        print(f"  {i+1}: {name} - {desc} (GUID: {guid})")

    print("\n--- Adapter Selection ---")
    print("You can apply interface-specific TCP tweaks (like disabling Nagle, immediate ACKs) to one or more adapters.")
    print("It's generally recommended for your primary physical gaming adapters (Ethernet, Wi-Fi).")
    print("It's generally NOT recommended for virtual adapters (like VPNs, e.g., WireSock) unless you're sure.")

    selected_guids = []
    while True:
        try:
            if not selected_guids:
                prompt_message = "\nEnter the number of the adapter you want to apply interface-specific tweaks to (or type 'skip' to not apply to any specific interface): "
            else:
                prompt_message = f"\nSelected GUIDs: {', '.join(selected_guids)}\nEnter the number of another adapter, or type 'done' if finished: "
            
            choice_str = input(prompt_message).strip().lower()

            if choice_str == 'done':
                if selected_guids:
                    break
                else: # 'done' typed without any selection, equivalent to skipping
                    print("No specific adapters selected for interface-specific tweaks.")
                    return [] 
            if choice_str == 'skip':
                print("Skipping interface-specific tweaks for all adapters.")
                return [] # Return empty list to signify skipping

            choice_num = int(choice_str)
            if 1 <= choice_num <= len(adapters):
                chosen_adapter = adapters[choice_num - 1]
                chosen_guid = chosen_adapter['InterfaceGuid']
                if chosen_guid not in selected_guids:
                    selected_guids.append(chosen_guid)
                    print(f"  Added: '{chosen_adapter['Name']}' - {chosen_guid}")
                else:
                    print(f"  Adapter '{chosen_adapter['Name']}' already selected.")
                
                if len(selected_guids) == len(adapters):
                    print("All available adapters with GUIDs have been selected.")
                    break
                
                if not selected_guids: # If first selection, ask to add more immediately
                     add_more_choice = input("  Do you want to add another adapter? (yes/no): ").strip().lower()
                     if add_more_choice not in ['yes', 'y']:
                        break
            else:
                print(f"  Invalid selection. Please enter a number between 1 and {len(adapters)}.")
        except ValueError:
            print("  Invalid input. Please enter a number, 'skip', or 'done'.")
            
    return selected_guids

def run_script():
    # Determine the directory of the current Python script
    try:
        current_script_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
    except NameError:
        current_script_dir = os.getcwd()
    except Exception:
        current_script_dir = os.getcwd()

    target_ps_script_path = os.path.join(current_script_dir, TARGET_PS_SCRIPT_NAME)

    if not os.path.exists(target_ps_script_path):
        message = (
            f"Error: The target PowerShell script '{TARGET_PS_SCRIPT_NAME}' was not found.\n"
            f"Looked in directory: '{current_script_dir}'\n\n"
            f"Please ensure '{TARGET_PS_SCRIPT_NAME}' is in the same folder as this Python script."
        )
        ctypes.windll.user32.MessageBoxW(None, message, "Script Not Found", 0x10 | 0x0) 
        print(message) 
        return 

    # --- New: Get adapters and select GUIDs ---
    adapters = get_network_adapters_from_powershell()
    if adapters is None: # Error occurred during adapter fetching
        print("Could not retrieve adapter list. Aborting launch of optimization script.")
        return
        
    selected_guids = select_guids_for_tweaks(adapters)
    guids_csv_for_ps = ",".join(selected_guids) # Create a comma-separated string of GUIDs

    print(f"\nAttempting to run '{target_ps_script_path}' as administrator...")
    if selected_guids:
        print(f"Interface-specific tweaks will be attempted for GUID(s): {guids_csv_for_ps}")
    else:
        print("No specific interface GUIDs selected; interface-specific tweaks section in PowerShell script will be skipped.")

    executable_to_run = "powershell.exe"
    # Pass the selected GUIDs CSV string as a parameter to the PowerShell script
    script_parameters = f'-NoProfile -ExecutionPolicy Bypass -File "{target_ps_script_path}" -TargetGuidsCsv "{guids_csv_for_ps}"'

    try:
        ret = ctypes.windll.shell32.ShellExecuteW(
            None, "runas", executable_to_run, script_parameters, None, 1
        )

        if ret <= 32:
            error_message = f"ShellExecuteW failed to start the script. Error code: {ret}\n\n"
            # ... (keep your existing detailed error code messages) ...
            if ret == 0: error_message += "The operating system is out of memory or resources."
            elif ret == 2: error_message += "File not found (powershell.exe or script path issue)."
            elif ret == 3: error_message += "Path not found."
            elif ret == 5: error_message += "Access denied (UAC prompt possibly denied or other permission issue)."
            elif ret == 1223: error_message += "The operation was canceled by the user (UAC prompt denied or closed)." # Common UAC denial
            else: error_message += "Refer to ShellExecuteW documentation for other error codes."

            ctypes.windll.user32.MessageBoxW(None, error_message, "Launch Error", 0x10 | 0x0)
            print(error_message)
        else:
            success_message = (
                f"Successfully requested to run '{TARGET_PS_SCRIPT_NAME}' as administrator.\n"
                "If UAC is enabled, a prompt should appear.\n"
                "The PowerShell script will run in a new window."
            )
            print(success_message)
            
    except Exception as e:
        exception_message = f"An exception occurred while trying to use ShellExecuteW: {e}"
        ctypes.windll.user32.MessageBoxW(None, exception_message, "Python Script Exception", 0x10 | 0x0)
        print(exception_message)

if __name__ == "__main__":
    run_script()
    print("\nThis Python script has finished its task of attempting to launch the PowerShell script.")
    if sys.stdout.isatty(): 
        input("Press Enter to close this Python script window...")