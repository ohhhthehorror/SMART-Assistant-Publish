"""
Auto-Update Module for SMART Assistant
Checks GitHub for new versions and handles updates
"""

import requests
import json
import os
import sys
import subprocess
import tempfile
from tkinter import messagebox
import threading

class AutoUpdater:
    def __init__(self, current_version, github_user, github_repo):
        """
        Initialize auto-updater
        
        Args:
            current_version: Current app version (e.g., "0.2.1")
            github_user: GitHub username
            github_repo: Repository name
        """
        self.current_version = current_version
        self.github_user = github_user
        self.github_repo = github_repo
        self.base_url = f"https://api.github.com/repos/{github_user}/{github_repo}"
        
        # Theme colors (will be set by app)
        self.is_dark_mode = True
        self.bg_card = "#1a1a1a"
        self.text_light = "#ffffff"
        self.accent = "#00d4ff"
        
    def check_for_updates(self, show_no_update_message=False):
        """
        Check if a new version is available
        
        Args:
            show_no_update_message: Show message if no update available
            
        Returns:
            dict with update info or None if no update
        """
        print(f"[AUTO-UPDATE] Checking for updates...")
        print(f"[AUTO-UPDATE] Current version: {self.current_version}")
        print(f"[AUTO-UPDATE] Checking: {self.base_url}/releases/latest")
        
        try:
            # Get latest release from GitHub
            response = requests.get(f"{self.base_url}/releases/latest", timeout=5)
            
            print(f"[AUTO-UPDATE] Response status: {response.status_code}")
            
            if response.status_code == 200:
                release_data = response.json()
                latest_version = release_data['tag_name'].replace('v', '')
                
                print(f"[AUTO-UPDATE] Latest version on GitHub: {latest_version}")
                
                # Compare versions
                if self.is_newer_version(latest_version, self.current_version):
                    print(f"[AUTO-UPDATE] Update available! {self.current_version} → {latest_version}")
                    
                    # Find the EXE asset
                    exe_url = None
                    for asset in release_data.get('assets', []):
                        if asset['name'].endswith('.exe'):
                            exe_url = asset['browser_download_url']
                            print(f"[AUTO-UPDATE] EXE found: {exe_url}")
                            break
                    
                    if not exe_url:
                        print(f"[AUTO-UPDATE] WARNING: No EXE file attached to release!")
                    
                    return {
                        'available': True,
                        'version': latest_version,
                        'download_url': exe_url,
                        'release_notes': release_data.get('body', 'No release notes available'),
                        'published_at': release_data.get('published_at', '')
                    }
                else:
                    print(f"[AUTO-UPDATE] Already on latest version: {self.current_version}")
                    if show_no_update_message:
                        messagebox.showinfo("No Updates", 
                                          f"You are running the latest version ({self.current_version})")
                    return None
            else:
                print(f"[AUTO-UPDATE] ERROR: HTTP {response.status_code}")
                print(f"[AUTO-UPDATE] Response: {response.text[:200]}")
                if show_no_update_message:
                    messagebox.showwarning("Update Check Failed", 
                                         "Could not check for updates. Please check your internet connection.")
                return None
                
        except Exception as e:
            print(f"[AUTO-UPDATE] EXCEPTION: {str(e)}")
            import traceback
            traceback.print_exc()
            
            if show_no_update_message:
                messagebox.showwarning("Update Check Failed", 
                                     f"Could not check for updates:\n{str(e)}")
            return None
    
    def is_newer_version(self, latest, current):
        """Compare version numbers"""
        try:
            latest_parts = [int(x) for x in latest.split('.')]
            current_parts = [int(x) for x in current.split('.')]
            
            # Pad shorter version with zeros
            while len(latest_parts) < len(current_parts):
                latest_parts.append(0)
            while len(current_parts) < len(latest_parts):
                current_parts.append(0)
            
            return latest_parts > current_parts
        except:
            return False
    
    def prompt_update(self, update_info):
        """
        Show update prompt to user with app theme
        
        Args:
            update_info: Dict with update information
            
        Returns:
            True if user wants to update, False otherwise
        """
        import tkinter as tk
        from tkinter import messagebox
        
        try:
            message = f"""A new version of SMART Assistant is available!

Current Version: {self.current_version}
Latest Version: {update_info['version']}

Release Notes:
{update_info['release_notes'][:300]}{'...' if len(update_info['release_notes']) > 300 else ''}

Would you like to download and install the update?"""
            
            # Create custom themed dialog
            dialog = tk.Toplevel()
            dialog.title("Update Available")
            dialog.configure(bg=self.bg_card)
            dialog.transient()
            dialog.grab_set()
            
            # Dark title bar
            if self.is_dark_mode:
                try:
                    dialog.update()
                    import ctypes
                    HWND = ctypes.windll.user32.GetParent(dialog.winfo_id())
                    DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                    ctypes.windll.dwmapi.DwmSetWindowAttribute(
                        HWND, DWMWA_USE_IMMERSIVE_DARK_MODE,
                        ctypes.byref(ctypes.c_int(1)), ctypes.sizeof(ctypes.c_int(1))
                    )
                except:
                    pass
            
            # Size and position
            dialog.geometry("550x450")
            dialog.update_idletasks()
            dialog.geometry(f"550x450+{dialog.winfo_screenwidth()//2 - 275}+{dialog.winfo_screenheight()//2 - 225}")
            
            # Content
            tk.Label(dialog, text=message, 
                    font=("Segoe UI", 11),
                    bg=self.bg_card,
                    fg=self.text_light,
                    wraplength=500,
                    justify="left").pack(pady=30, padx=30)
            
            result = {"value": False}
            
            def on_yes():
                result["value"] = True
                dialog.destroy()
            
            def on_no():
                dialog.destroy()
            
            # Buttons
            btn_frame = tk.Frame(dialog, bg=self.bg_card)
            btn_frame.pack(pady=(0, 20))
            
            tk.Button(btn_frame, 
                     text="Yes, Update", 
                     command=on_yes, 
                     width=15,
                     bg=self.accent, 
                     fg="white",
                     font=("Segoe UI", 10, "bold"), 
                     relief="flat", 
                     cursor="hand2",
                     padx=15,
                     pady=8).pack(side="left", padx=8)
            
            tk.Button(btn_frame, 
                     text="Not Now", 
                     command=on_no, 
                     width=15,
                     bg="#999999", 
                     fg="white",
                     font=("Segoe UI", 10, "bold"), 
                     relief="flat", 
                     cursor="hand2",
                     padx=10,
                     pady=6).pack(side="left", padx=8)
            
            dialog.wait_window()
            return result["value"]
            
        except Exception as e:
            print(f"Error creating update dialog: {str(e)}")
            import traceback
            traceback.print_exc()
            # Fallback to simple messagebox
            return messagebox.askyesno("Update Available", 
                                      f"New version {update_info['version']} available. Update now?")
    
    def download_and_install(self, download_url, progress_callback=None):
        """
        Download new version and install it
        
        Args:
            download_url: URL to download EXE
            progress_callback: Function to call with progress updates
        """
        try:
            if progress_callback:
                progress_callback("Downloading update...")
            
            # Download the new EXE
            response = requests.get(download_url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            
            # Save to temp file first
            temp_dir = tempfile.gettempdir()
            new_exe_path = os.path.join(temp_dir, "SMART_Assistant_NEW.exe")
            
            downloaded = 0
            with open(new_exe_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size > 0:
                            progress = int((downloaded / total_size) * 100)
                            progress_callback(f"Downloading update... {progress}%")
            
            if progress_callback:
                progress_callback("Installing update...")
            
            # Create update script
            # Get the REAL exe path (not the temp extraction folder)
            if getattr(sys, 'frozen', False):
                # Running as EXE
                if hasattr(sys, '_MEIPASS'):
                    # PyInstaller onefile mode - sys.executable points to temp folder
                    # We need to find the actual EXE location
                    # Check if there's a way to get it from environment or command line
                    import os
                    # Try to get from argv[0] which might have the real path
                    if len(sys.argv) > 0 and os.path.exists(sys.argv[0]):
                        current_exe = os.path.abspath(sys.argv[0])
                    else:
                        # Fallback to sys.executable
                        current_exe = sys.executable
                else:
                    current_exe = sys.executable
            else:
                current_exe = "SMART_Assistant.exe"
            
            current_dir = os.path.dirname(os.path.abspath(current_exe))
            exe_name = os.path.basename(current_exe)
            
            # Create batch script to replace the exe
            update_script = f"""@echo off
title SMART Assistant Update
echo.
echo ============================================
echo SMART Assistant - Installing Update
echo ============================================
echo.
echo Current directory: {current_dir}
echo EXE name: {exe_name}
echo.
echo Waiting for application to close...
timeout /t 3 /nobreak >nul

:WAIT_LOOP
tasklist /FI "IMAGENAME eq {exe_name}" 2>NUL | find /I /N "{exe_name}">NUL
if "%ERRORLEVEL%"=="0" (
    timeout /t 1 /nobreak >nul
    goto WAIT_LOOP
)

echo.
echo Replacing executable...
cd /d "{current_dir}"

if exist "{exe_name}" (
    del /F /Q "{exe_name}.old" >nul 2>&1
    move /Y "{exe_name}" "{exe_name}.old" >nul 2>&1
)

move /Y "{new_exe_path}" "{exe_name}" >nul 2>&1

if exist "{exe_name}.old" (
    del /F /Q "{exe_name}.old" >nul 2>&1
)

if exist "{exe_name}" (
    echo.
    echo ============================================
    echo Update Complete! Starting new version...
    echo ============================================
    echo.
    timeout /t 1 /nobreak >nul
    
    REM Start the new version from the correct directory
    cd /d "{current_dir}"
    start "" "{current_dir}\{exe_name}"
    
    timeout /t 1 /nobreak >nul
    echo New version launched successfully!
) else (
    echo.
    echo ERROR: Update failed!
    echo The new executable was not found.
    echo Please download manually from GitHub
    pause
)

REM Clean up this script
timeout /t 2 /nobreak >nul
del "%~f0" >nul 2>&1
exit
"""
            
            update_script_path = os.path.join(current_dir, "update_smart_assistant.bat")
            with open(update_script_path, 'w') as f:
                f.write(update_script)
            
            if progress_callback:
                progress_callback("Update ready! Restarting...")
            
            # Close the progress window if it exists
            try:
                import tkinter as tk
                for widget in tk.Tk().winfo_children():
                    if isinstance(widget, tk.Toplevel):
                        widget.destroy()
            except:
                pass
            
            # Launch update script (show console for visibility)
            if os.name == 'nt':
                # Windows: Start in new console window
                subprocess.Popen(['cmd', '/c', 'start', 'cmd', '/c', update_script_path])
            else:
                # Linux/Mac
                subprocess.Popen(update_script_path, shell=True)
            
            # Small delay to ensure script starts
            import time
            time.sleep(0.5)
            
            # Force exit the current application
            try:
                # Try graceful exit first
                import tkinter as tk
                root = tk._default_root
                if root:
                    root.quit()
                    root.destroy()
            except:
                pass
            
            # Hard exit as last resort
            os._exit(0)
            
        except Exception as e:
            if progress_callback:
                progress_callback(f"Update failed: {str(e)}")
            messagebox.showerror("Update Failed", 
                               f"Could not install update:\n{str(e)}\n\nPlease download manually from GitHub.")
            return False
    
    def check_and_prompt_update(self, parent_window=None):
        """
        Check for updates and prompt user (runs in background thread)
        
        Args:
            parent_window: Parent Tkinter window (for proper dialog positioning)
        """
        print("[AUTO-UPDATE] check_and_prompt_update called")
        
        def check_thread():
            print("[AUTO-UPDATE] Thread started")
            try:
                update_info = self.check_for_updates(show_no_update_message=False)
                
                print(f"[AUTO-UPDATE] Update info: {update_info}")
                
                if update_info and update_info['available']:
                    print("[AUTO-UPDATE] Update available, showing prompt...")
                    # Show prompt on main thread
                    if parent_window:
                        parent_window.after(0, lambda: self._show_update_prompt(update_info))
                    else:
                        self._show_update_prompt(update_info)
                else:
                    print("[AUTO-UPDATE] No update available")
                    
            except Exception as e:
                print(f"[AUTO-UPDATE] EXCEPTION in check_thread: {str(e)}")
                import traceback
                traceback.print_exc()
        
        thread = threading.Thread(target=check_thread, daemon=True)
        thread.start()
        print("[AUTO-UPDATE] Thread launched")
    
    def _show_update_prompt(self, update_info):
        """Show update prompt and handle download"""
        if self.prompt_update(update_info):
            if update_info.get('download_url'):
                # Show progress window
                progress_window = self._create_progress_window()
                
                def download_thread():
                    self.download_and_install(
                        update_info['download_url'],
                        progress_callback=lambda msg: self._update_progress(progress_window, msg)
                    )
                
                thread = threading.Thread(target=download_thread, daemon=True)
                thread.start()
            else:
                messagebox.showwarning("Update Error", 
                                     "Could not find download link. Please visit GitHub manually.")
    
    def _create_progress_window(self):
        """Create a simple progress window"""
        import tkinter as tk
        
        window = tk.Toplevel()
        window.title("Updating SMART Assistant")
        window.geometry("400x150")
        window.resizable(False, False)
        
        # Center window
        window.update_idletasks()
        x = (window.winfo_screenwidth() // 2) - (400 // 2)
        y = (window.winfo_screenheight() // 2) - (150 // 2)
        window.geometry(f"400x150+{x}+{y}")
        
        label = tk.Label(window, text="Preparing update...", 
                        font=("Segoe UI", 11), pady=30)
        label.pack()
        
        progress_label = tk.Label(window, text="", font=("Segoe UI", 9))
        progress_label.pack()
        
        window.label = label
        window.progress_label = progress_label
        
        return window
    
    def _update_progress(self, window, message):
        """Update progress window"""
        if window and window.winfo_exists():
            window.label.config(text=message)
            window.update()
