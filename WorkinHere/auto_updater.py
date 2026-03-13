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
import logging

logger = logging.getLogger(__name__)

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
        logger.info("Checking for updates...")
        logger.info(f"Current version: {self.current_version}")
        logger.info(f"Checking: {self.base_url}/releases/latest")
        print(f"[AUTO-UPDATE] Checking for updates...")
        print(f"[AUTO-UPDATE] Current version: {self.current_version}")
        print(f"[AUTO-UPDATE] Checking: {self.base_url}/releases/latest")
        
        try:
            # Get latest release from GitHub
            response = requests.get(f"{self.base_url}/releases/latest", timeout=5)
            
            logger.info(f"Response status: {response.status_code}")
            print(f"[AUTO-UPDATE] Response status: {response.status_code}")
            
            if response.status_code == 200:
                release_data = response.json()
                latest_version = release_data['tag_name'].replace('v', '')
                
                logger.info(f"Latest version on GitHub: {latest_version}")
                print(f"[AUTO-UPDATE] Latest version on GitHub: {latest_version}")
                
                # Compare versions
                if self.is_newer_version(latest_version, self.current_version):
                    logger.info(f"Update available! {self.current_version} → {latest_version}")
                    print(f"[AUTO-UPDATE] Update available! {self.current_version} → {latest_version}")
                    
                    # Find the EXE asset
                    exe_url = None
                    for asset in release_data.get('assets', []):
                        if asset['name'].endswith('.exe'):
                            exe_url = asset['browser_download_url']
                            logger.info(f"EXE found: {exe_url}")
                            print(f"[AUTO-UPDATE] EXE found: {exe_url}")
                            break
                    
                    if not exe_url:
                        logger.warning("No EXE file attached to release!")
                        print(f"[AUTO-UPDATE] WARNING: No EXE file attached to release!")
                    
                    return {
                        'available': True,
                        'version': latest_version,
                        'download_url': exe_url,
                        'release_notes': release_data.get('body', 'No release notes available'),
                        'published_at': release_data.get('published_at', '')
                    }
                else:
                    logger.info(f"Already on latest version: {self.current_version}")
                    print(f"[AUTO-UPDATE] Already on latest version: {self.current_version}")
                    if show_no_update_message:
                        messagebox.showinfo("No Updates", 
                                          f"You are running the latest version ({self.current_version})")
                    return None
            else:
                logger.error(f"HTTP {response.status_code}")
                logger.error(f"Response: {response.text[:200]}")
                print(f"[AUTO-UPDATE] ERROR: HTTP {response.status_code}")
                print(f"[AUTO-UPDATE] Response: {response.text[:200]}")
                if show_no_update_message:
                    messagebox.showwarning("Update Check Failed", 
                                         "Could not check for updates. Please check your internet connection.")
                return None
                
        except Exception as e:
            logger.exception(f"Exception during update check: {str(e)}")
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
            
            result = latest_parts > current_parts
            
            logger.info(f"Version comparison:")
            logger.info(f"  Current: {current} = {current_parts}")
            logger.info(f"  Latest:  {latest} = {latest_parts}")
            logger.info(f"  Is newer? {result}")
            
            return result
        except Exception as e:
            logger.error(f"Version comparison error: {e}")
            return False
    
    def prompt_update(self, update_info, parent_window=None):
        """
        Show update prompt to user with app theme
        
        Args:
            update_info: Dict with update information
            parent_window: Parent Tkinter window for positioning
            
        Returns:
            True if user wants to update, False otherwise
        """
        import tkinter as tk
        from tkinter import scrolledtext
        
        try:
            # Create custom themed dialog
            dialog = tk.Toplevel()
            dialog.title("Update Available")
            dialog.configure(bg=self.bg_card)
            
            # Make dialog appear on same screen as parent
            if parent_window:
                dialog.transient(parent_window)
            
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
            
            # Size and position relative to parent window
            dialog.geometry("600x500")
            dialog.update_idletasks()
            
            if parent_window:
                # Position relative to parent window
                parent_x = parent_window.winfo_x()
                parent_y = parent_window.winfo_y()
                parent_width = parent_window.winfo_width()
                parent_height = parent_window.winfo_height()
                
                x = parent_x + (parent_width // 2) - 300
                y = parent_y + (parent_height // 2) - 250
            else:
                # Fallback to screen center
                x = dialog.winfo_screenwidth()//2 - 300
                y = dialog.winfo_screenheight()//2 - 250
            
            dialog.geometry(f"600x500+{x}+{y}")
            
            # Header
            header_frame = tk.Frame(dialog, bg=self.bg_card)
            header_frame.pack(fill="x", padx=30, pady=(30, 10))
            
            tk.Label(header_frame, 
                    text="🔔 Update Available", 
                    font=("Segoe UI", 16, "bold"),
                    bg=self.bg_card,
                    fg=self.accent).pack(anchor="w")
            
            # Version info
            version_frame = tk.Frame(dialog, bg=self.bg_card)
            version_frame.pack(fill="x", padx=30, pady=(0, 10))
            
            tk.Label(version_frame,
                    text=f"Current Version: {self.current_version}",
                    font=("Segoe UI", 11),
                    bg=self.bg_card,
                    fg=self.text_light).pack(anchor="w")
            
            tk.Label(version_frame,
                    text=f"Latest Version: {update_info['version']}",
                    font=("Segoe UI", 11, "bold"),
                    bg=self.bg_card,
                    fg=self.accent).pack(anchor="w")
            
            # Release notes label
            tk.Label(dialog,
                    text="Release Notes:",
                    font=("Segoe UI", 11, "bold"),
                    bg=self.bg_card,
                    fg=self.text_light).pack(anchor="w", padx=30, pady=(10, 5))
            
            # Scrollable release notes
            notes_frame = tk.Frame(dialog, bg=self.bg_card, height=200)
            notes_frame.pack(fill="x", padx=30, pady=(0, 10))
            notes_frame.pack_propagate(False)  # Prevent frame from shrinking
            
            notes_text = scrolledtext.ScrolledText(
                notes_frame,
                wrap=tk.WORD,
                height=10,  # Explicit height in lines
                font=("Segoe UI", 10),
                bg="#2a2a2a" if self.is_dark_mode else "#f5f5f5",
                fg=self.text_light,
                relief="solid",
                borderwidth=1,
                padx=10,
                pady=10
            )
            notes_text.pack(fill="both", expand=True)
            notes_text.insert("1.0", update_info['release_notes'])
            notes_text.config(state="disabled")  # Make read-only
            
            # Question
            tk.Label(dialog,
                    text="Would you like to download and install the update?",
                    font=("Segoe UI", 11),
                    bg=self.bg_card,
                    fg=self.text_light).pack(pady=(10, 15))
            
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
        Download new version and replace the current EXE
        
        Args:
            download_url: URL to download EXE
            progress_callback: Function to call with progress updates
        """
        try:
            logger.info(f"Starting download from: {download_url}")
            if progress_callback:
                progress_callback("Downloading update...")
            
            # Download the new EXE
            response = requests.get(download_url, stream=True)
            total_size = int(response.headers.get('content-length', 0))
            logger.info(f"Download size: {total_size} bytes")
            
            # Get current EXE location
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
            else:
                current_exe = os.path.abspath("SMART_Assistant.exe")
            
            current_dir = os.path.dirname(current_exe)
            exe_name = os.path.basename(current_exe)
            logger.info(f"Current EXE: {current_exe}")
            logger.info(f"Current directory: {current_dir}")
            
            # Download to temp location
            temp_dir = tempfile.gettempdir()
            new_exe_path = os.path.join(temp_dir, "SMART_Assistant_NEW.exe")
            logger.info(f"Downloading to: {new_exe_path}")
            
            downloaded = 0
            with open(new_exe_path, 'wb') as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        downloaded += len(chunk)
                        if progress_callback and total_size > 0:
                            progress = int((downloaded / total_size) * 100)
                            progress_callback(f"Downloading update... {progress}%")
            
            logger.info(f"Download complete: {downloaded} bytes")
            if progress_callback:
                progress_callback("Installing update...")
            
            # Create simple update script
            update_script = f"""@echo off
echo Waiting for SMART Assistant to close...
timeout /t 2 /nobreak >nul

cd /d "{current_dir}"

if exist "{exe_name}.old" (
    del /F /Q "{exe_name}.old" >nul 2>&1
)

if exist "{exe_name}" (
    move /Y "{exe_name}" "{exe_name}.old" >nul 2>&1
)

move /Y "{new_exe_path}" "{exe_name}" >nul 2>&1

if exist "{exe_name}.old" (
    del /F /Q "{exe_name}.old" >nul 2>&1
)

del "%~f0" >nul 2>&1
exit
"""
            
            update_script_path = os.path.join(current_dir, "update_smart_assistant.bat")
            with open(update_script_path, 'w') as f:
                f.write(update_script)
            
            # Launch update script
            if os.name == 'nt':
                subprocess.Popen(['cmd', '/c', update_script_path], 
                               creationflags=subprocess.CREATE_NO_WINDOW)
            
            # Show completion message
            import tkinter as tk
            from tkinter import messagebox
            
            try:
                # Create themed dialog
                dialog = tk.Toplevel()
                dialog.title("Update Complete")
                dialog.configure(bg=self.bg_card)
                
                # Position relative to parent window
                if hasattr(self, 'parent_window') and self.parent_window:
                    dialog.transient(self.parent_window)
                
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
                
                dialog.geometry("500x250")
                dialog.update_idletasks()
                
                # Position relative to parent window
                if hasattr(self, 'parent_window') and self.parent_window:
                    parent_x = self.parent_window.winfo_x()
                    parent_y = self.parent_window.winfo_y()
                    parent_width = self.parent_window.winfo_width()
                    parent_height = self.parent_window.winfo_height()
                    
                    x = parent_x + (parent_width // 2) - 250
                    y = parent_y + (parent_height // 2) - 125
                else:
                    x = dialog.winfo_screenwidth()//2 - 250
                    y = dialog.winfo_screenheight()//2 - 125
                
                dialog.geometry(f"500x250+{x}+{y}")
                
                # Message
                msg = """✅ Update Downloaded Successfully!

The update will be installed when you close this application.

Please close SMART Assistant now and reopen it to use the new version."""
                
                tk.Label(dialog, text=msg, 
                        font=("Segoe UI", 11),
                        bg=self.bg_card,
                        fg=self.text_light,
                        wraplength=450,
                        justify="left").pack(pady=30, padx=30)
                
                def on_ok():
                    dialog.destroy()
                    # Exit the app
                    os._exit(0)
                
                tk.Button(dialog, 
                         text="OK - Close App", 
                         command=on_ok, 
                         width=20,
                         bg=self.accent, 
                         fg="white",
                         font=("Segoe UI", 10, "bold"), 
                         relief="flat", 
                         cursor="hand2",
                         padx=15,
                         pady=8).pack(pady=(0, 20))
                
                dialog.wait_window()
                
            except:
                # Fallback to simple messagebox
                messagebox.showinfo("Update Complete", 
                                   "Update downloaded! Please close and reopen SMART Assistant.")
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
                        parent_window.after(0, lambda: self._show_update_prompt(update_info, parent_window))
                    else:
                        self._show_update_prompt(update_info, None)
                else:
                    print("[AUTO-UPDATE] No update available")
                    
            except Exception as e:
                print(f"[AUTO-UPDATE] EXCEPTION in check_thread: {str(e)}")
                import traceback
                traceback.print_exc()
        
        thread = threading.Thread(target=check_thread, daemon=True)
        thread.start()
        print("[AUTO-UPDATE] Thread launched")
    
    def _show_update_prompt(self, update_info, parent_window=None):
        """Show update prompt and handle download"""
        # Store parent for use in download dialogs
        self.parent_window = parent_window
        
        if self.prompt_update(update_info, parent_window):
            if update_info.get('download_url'):
                # Show progress window
                progress_window = self._create_progress_window(parent_window)
                
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
    
    def _create_progress_window(self, parent_window=None):
        """Create a themed progress window with progress bar"""
        import tkinter as tk
        
        window = tk.Toplevel()
        window.title("Updating SMART Assistant")
        window.configure(bg=self.bg_card)
        window.geometry("500x220")
        window.resizable(False, False)
        
        # Make window appear on same screen as parent
        if parent_window:
            window.transient(parent_window)
        
        # Dark title bar if in dark mode
        if self.is_dark_mode:
            try:
                window.update()
                import ctypes
                HWND = ctypes.windll.user32.GetParent(window.winfo_id())
                DWMWA_USE_IMMERSIVE_DARK_MODE = 20
                ctypes.windll.dwmapi.DwmSetWindowAttribute(
                    HWND, DWMWA_USE_IMMERSIVE_DARK_MODE,
                    ctypes.byref(ctypes.c_int(1)), ctypes.sizeof(ctypes.c_int(1))
                )
            except:
                pass
        
        # Position relative to parent window
        window.update_idletasks()
        
        if parent_window:
            parent_x = parent_window.winfo_x()
            parent_y = parent_window.winfo_y()
            parent_width = parent_window.winfo_width()
            parent_height = parent_window.winfo_height()
            
            x = parent_x + (parent_width // 2) - 250
            y = parent_y + (parent_height // 2) - 110
        else:
            x = (window.winfo_screenwidth() // 2) - 250
            y = (window.winfo_screenheight() // 2) - 110
        
        window.geometry(f"500x220+{x}+{y}")
        
        # Main label
        label = tk.Label(window, text="Preparing update...", 
                        font=("Segoe UI", 12, "bold"),
                        bg=self.bg_card,
                        fg=self.text_light,
                        pady=20)
        label.pack()
        
        # Progress bar container
        progress_container = tk.Frame(window, bg=self.bg_card)
        progress_container.pack(fill="x", padx=40, pady=10)
        
        # Progress bar background
        progress_bg = tk.Frame(progress_container, bg="#333333" if self.is_dark_mode else "#e0e0e0", height=30)
        progress_bg.pack(fill="x")
        
        # Progress bar fill
        progress_fill = tk.Frame(progress_bg, bg=self.accent, height=30)
        progress_fill.place(x=0, y=0, relwidth=0, relheight=1)
        
        # Progress percentage label
        progress_label = tk.Label(window, text="0%", 
                                 font=("Segoe UI", 10),
                                 bg=self.bg_card,
                                 fg=self.text_light)
        progress_label.pack(pady=10)
        
        window.label = label
        window.progress_label = progress_label
        window.progress_fill = progress_fill
        
        return window
    
    def _update_progress(self, window, message):
        """Update progress window with percentage"""
        if window and window.winfo_exists():
            window.label.config(text=message)
            
            # Extract percentage from message if present
            import re
            match = re.search(r'(\d+)%', message)
            if match:
                percentage = int(match.group(1))
                window.progress_fill.place(relwidth=percentage/100)
                window.progress_label.config(text=f"{percentage}%")
            
            window.update()
