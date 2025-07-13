import os
import sys
import tkinter as tk
from tkinter import messagebox
import ctypes
import subprocess
import time
import threading
import keyboard
import winreg
import getpass


def create_startup_entry():
    """Create a registry entry to run the application at startup"""
    try:
        # Get the full path of the current script
        script_path = os.path.abspath(sys.argv[0])

        # If running from .py file, use pythonw.exe for background execution
        if script_path.endswith('.py'):
            exe_path = f'"{sys.executable}w" "{script_path}"'
        else:
            # If it's an exe, just use the exe path
            exe_path = f'"{script_path}"'

        # Open the run key
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE
        )

        # Set the value
        winreg.SetValueEx(
            key,
            "GPUSwitcher",
            0,
            winreg.REG_SZ,
            exe_path
        )

        # Close the key
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error creating startup entry: {e}")
        return False


def remove_startup_entry():
    """Remove the registry entry for startup"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE
        )

        try:
            winreg.DeleteValue(key, "GPUSwitcher")
        except FileNotFoundError:
            # Key doesn't exist, that's fine
            pass

        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error removing startup entry: {e}")
        return False


def check_startup_entry():
    """Check if the startup entry exists"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_READ
        )

        try:
            winreg.QueryValueEx(key, "GPUSwitcher")
            exists = True
        except FileNotFoundError:
            exists = False

        winreg.CloseKey(key)
        return exists
    except Exception:
        return False


def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


# Restart with admin rights if needed
if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 0)
    sys.exit()


def execute_command(command, visible=False):
    """Execute command with specified visibility"""
    if visible:
        return subprocess.run(command, shell=True, text=True)
    else:
        startupinfo = subprocess.STARTUPINFO()
        startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
        startupinfo.wShowWindow = 0  # SW_HIDE
        return subprocess.run(command, shell=True, text=True,
                              startupinfo=startupinfo,
                              stdout=subprocess.PIPE,
                              stderr=subprocess.PIPE)


def find_nvidia_device_id():
    """Find NVIDIA GPU device ID automatically"""
    result = execute_command('pnputil /enum-devices /class Display')
    lines = result.stdout.splitlines()

    for i, line in enumerate(lines):
        if 'NVIDIA' in line and 'Description:' in line:
            # Look for the Instance ID in nearby lines
            for j in range(max(0, i - 5), min(i + 5, len(lines))):
                if 'Instance ID:' in lines[j]:
                    return lines[j].split(':', 1)[1].strip()
    return None


def switch_to_eco_mode(status_var=None):
    """Switch to Eco Mode (iGPU + 60Hz)"""
    if status_var:
        status_var.set("Switching to Eco Mode...")
        root.update_idletasks()

    # Find NVIDIA device ID
    device_id = find_nvidia_device_id()
    if not device_id:
        if status_var:
            status_var.set("Error: Could not find NVIDIA device")
        return

    # Disable NVIDIA GPU silently
    execute_command(f'pnputil /disable-device "{device_id}"')

    # Force hardware changes silently
    execute_command('pnputil /scan-devices')

    # Allow system time to process changes
    time.sleep(2)

    # Change refresh rate silently
    execute_command('nircmd.exe setdisplay 1920 1080 32 60')

    if status_var:
        status_var.set("✓ Eco Mode Active (iGPU + 60Hz)")


def switch_to_standard_mode(status_var=None):
    """Switch to Standard Mode (NVIDIA + 144Hz)"""
    if status_var:
        status_var.set("Switching to Standard Mode...")
        root.update_idletasks()

    # Find NVIDIA device ID
    device_id = find_nvidia_device_id()
    if not device_id:
        if status_var:
            status_var.set("Error: Could not find NVIDIA device")
        return

    # Enable NVIDIA GPU silently
    execute_command(f'pnputil /enable-device "{device_id}"')

    # Force hardware changes silently
    execute_command('pnputil /scan-devices')

    # Allow more time for GPU initialization
    time.sleep(3)

    # Change refresh rate silently
    execute_command('nircmd.exe setdisplay 1920 1080 32 144')

    if status_var:
        status_var.set("✓ Standard Mode Active (NVIDIA + 144Hz)")


def run_in_thread(func, status_var=None):
    """Run the specified function in a separate thread to keep UI responsive"""
    threading.Thread(target=func, args=(status_var,), daemon=True).start()


# Global hotkey handlers
def eco_mode_global_hotkey():
    run_in_thread(switch_to_eco_mode)


def standard_mode_global_hotkey():
    run_in_thread(switch_to_standard_mode)


# Check if script is called with --background flag
if len(sys.argv) > 1 and sys.argv[1] == "--background":
    # Register global hotkeys in background mode
    keyboard.add_hotkey('ctrl+shift+h', eco_mode_global_hotkey)
    keyboard.add_hotkey('ctrl+shift+g', standard_mode_global_hotkey)

    # Keep script running
    keyboard.wait()
    sys.exit()

# Register global hotkeys for normal mode
keyboard.add_hotkey('ctrl+shift+h', eco_mode_global_hotkey)
keyboard.add_hotkey('ctrl+shift+g', standard_mode_global_hotkey)

# Create minimal GUI
root = tk.Tk()
root.title("GPU Switcher")
root.geometry("300x190")
root.resizable(False, False)

# Use a clean, professional style
root.configure(bg="#f0f0f0")
main_frame = tk.Frame(root, bg="#f0f0f0", padx=15, pady=15)
main_frame.pack(fill="both", expand=True)

# Status variable
status_var = tk.StringVar()
status_var.set("Ready")

# Create stylish buttons
eco_button = tk.Button(
    main_frame,
    text="Switch to Eco Mode",
    command=lambda: run_in_thread(switch_to_eco_mode, status_var),
    bg="#4CAF50",
    fg="white",
    width=20,
    height=1,
    font=("Segoe UI", 10)
)
eco_button.pack(pady=5)

standard_button = tk.Button(
    main_frame,
    text="Switch to Standard Mode",
    command=lambda: run_in_thread(switch_to_standard_mode, status_var),
    bg="#2196F3",
    fg="white",
    width=20,
    height=1,
    font=("Segoe UI", 10)
)
standard_button.pack(pady=5)

# Status label
status_label = tk.Label(
    main_frame,
    textvariable=status_var,
    bg="#f0f0f0",
    fg="#333333",
    font=("Segoe UI", 9)
)
status_label.pack(pady=5)

# Add shortcut information
shortcut_label = tk.Label(
    main_frame,
    text="Global Shortcuts: Ctrl+Shift+H (Eco) | Ctrl+Shift+G (Standard)",
    bg="#f0f0f0",
    fg="#666666",
    font=("Segoe UI", 8)
)
shortcut_label.pack()

# Startup checkbox
startup_var = tk.BooleanVar()
startup_var.set(check_startup_entry())


def toggle_startup():
    if startup_var.get():
        create_startup_entry()
    else:
        remove_startup_entry()


startup_check = tk.Checkbutton(
    main_frame,
    text="Run at startup (with hotkeys enabled)",
    variable=startup_var,
    command=toggle_startup,
    bg="#f0f0f0",
    fg="#333333",
    font=("Segoe UI", 8)
)
startup_check.pack(pady=5)


def on_closing():
    # Ask if user wants to keep the app running in the background
    if messagebox.askyesno("Minimize to Tray", "Do you want to keep the hotkeys active in the background?"):
        # Start a new instance in background mode
        subprocess.Popen([sys.executable, sys.argv[0], "--background"],
                         creationflags=subprocess.CREATE_NO_WINDOW)

    # Clean up keyboard hooks
    keyboard.unhook_all()
    root.destroy()


root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()