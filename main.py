import tkinter as tk
from tkinter import ttk, filedialog, messagebox, Menu
from app_data import AppData
from winotify import Notification, audio
import os
import base64
import subprocess
import sys
import ctypes
import json
import threading
import webbrowser
from io import BytesIO
try:
    from PIL import Image, ImageTk
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False

class PowerShellManager:
    def __init__(self, root):
        self.root = root
        self.root.title("PowerShell Script Manager")
        self.app_data = AppData()
        
        # Configure root window
        self.root.geometry("800x600")
        
        # Load button icons
        self.load_icons()
        
        # Set up system tray icon
        self.setup_system_tray()
        
        # Handle window close button
        self.root.protocol('WM_DELETE_WINDOW', self.hide_window)
        
        # Store PowerShell update statuses
        self.powershell_status = {}
        
        # Start background update check
        self.start_update_check()
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Create tabs
        self.home_tab = ttk.Frame(self.notebook)
        self.powershell_tab = ttk.Frame(self.notebook)
        self.folders_tab = ttk.Frame(self.notebook)
        
        self.notebook.add(self.home_tab, text='Scripts')
        self.notebook.add(self.powershell_tab, text='PowerShell')
        self.notebook.add(self.folders_tab, text='Settings')
        
        self.setup_home_tab()
        self.setup_powershell_tab()
        self.setup_folders_tab()
        
        # Initial refresh of scripts with startup notification
        self.refresh_script_list(show_startup_notification=True)
        
    def check_execution_policy(self):
        """Check if PowerShell execution policy allows scripts to run"""
        try:
            # Run PowerShell command to get execution policy
            result = subprocess.run(
                ['powershell.exe', '-Command', 'Get-ExecutionPolicy'],
                capture_output=True,
                text=True
            )
            
            policy = result.stdout.strip().lower()
            
            # Policies that allow script execution
            allowed_policies = ['unrestricted', 'remotesigned', 'bypass', 'allsigned']
            
            return policy in allowed_policies
        except Exception:
            # If we can't determine the policy, assume it's restricted
            return False

    def setup_home_tab(self):
        # Create main split between list and preview
        split_frame = ttk.PanedWindow(self.home_tab, orient=tk.HORIZONTAL)
        split_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Create frame for scripts list
        list_frame = ttk.Frame(split_frame)
        split_frame.add(list_frame, weight=1)
        
        # Create frame for script preview
        preview_frame = ttk.Frame(split_frame)
        split_frame.add(preview_frame, weight=1)
        
        # Check execution policy and show warning if needed
        if not self.check_execution_policy():
            warning_frame = ttk.Frame(self.home_tab)
            warning_frame.pack(side='top', fill='x', padx=5, pady=(0, 5))
            
            # Create a container for the warning label (left-aligned)
            left_container = ttk.Frame(warning_frame)
            left_container.pack(side='left', fill='x', expand=True)
            
            warning_label = tk.Label(
                left_container, 
                text="Running scripts is disabled on this system.",
                fg="red",
                justify=tk.LEFT
            )
            warning_label.pack(side='left', anchor='w')
            
            def open_docs():
                import webbrowser
                webbrowser.open("https://go.microsoft.com/fwlink/?LinkID=135170")
            
            # Create button container for right alignment
            button_container = ttk.Frame(warning_frame)
            button_container.pack(side='right')
            
            learn_more_btn = ttk.Button(
                button_container, 
                text="Learn More", 
                command=open_docs
            )
            learn_more_btn.pack(side='right', padx=(0, 5))
        
        # Create button frame at the top, aligned right
        self.button_frame = ttk.Frame(preview_frame)
        self.button_frame.pack(side='top', fill='x', padx=5, pady=(5,0))
        
        # Create a container for the buttons aligned to the right
        self.action_buttons_frame = ttk.Frame(self.button_frame)
        
        # Create action buttons but don't pack them initially
        if PILLOW_AVAILABLE and hasattr(self, 'icons'):
            self.refresh_btn = ttk.Button(self.action_buttons_frame, text="Refresh", 
                                         image=self.icons.get('refresh'), compound=tk.LEFT, 
                                         command=self.refresh_preview)
            self.notepad_btn = ttk.Button(self.action_buttons_frame, text="Open in Notepad", 
                                         image=self.icons.get('notepad'), compound=tk.LEFT,
                                         command=self.open_in_notepad)
            self.run_btn = ttk.Button(self.action_buttons_frame, text="Run", 
                                     image=self.icons.get('run'), compound=tk.LEFT,
                                     command=self.run_script)
            self.runas_btn = ttk.Button(self.action_buttons_frame, text="Run As...", 
                                       image=self.icons.get('run_as'), compound=tk.LEFT,
                                       command=self.run_script_as)
        else:
            self.refresh_btn = ttk.Button(self.action_buttons_frame, text="Refresh", 
                                         command=self.refresh_preview)
            self.notepad_btn = ttk.Button(self.action_buttons_frame, text="Open in Notepad", 
                                         command=self.open_in_notepad)
            self.run_btn = ttk.Button(self.action_buttons_frame, text="Run", 
                                     command=self.run_script)
            self.runas_btn = ttk.Button(self.action_buttons_frame, text="Run As...", 
                                       command=self.run_script_as)
        
        # Create a LabelFrame for the preview section
        self.preview_label_frame = ttk.LabelFrame(preview_frame, text="Script Preview")
        self.preview_label_frame.pack(side='top', fill='both', expand=True, padx=5, pady=(5,0))
        
        # Initially hide action buttons since no script is selected
        self.show_action_buttons(False)
        
        # Create container frame for text and scrollbars
        text_container = ttk.Frame(self.preview_label_frame)
        text_container.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Create text widget for script preview
        self.preview_text = tk.Text(text_container, wrap=tk.NONE, font=('Consolas', 10))
        self.preview_text.pack(side='left', fill='both', expand=True)
        
        # Add scrollbars
        y_scrollbar = ttk.Scrollbar(text_container, orient='vertical', command=self.preview_text.yview)
        y_scrollbar.pack(side='right', fill='y')
        
        x_scrollbar = ttk.Scrollbar(self.preview_label_frame, orient='horizontal', command=self.preview_text.xview)
        x_scrollbar.pack(side='bottom', fill='x')
        
        self.preview_text.configure(yscrollcommand=y_scrollbar.set, xscrollcommand=x_scrollbar.set)
        self.preview_text.configure(state='disabled')
        
        # Create top button frame
        button_frame = ttk.Frame(list_frame)
        button_frame.pack(side='top', fill='x', pady=5)
        
        # Add refresh button to the right with icon
        if PILLOW_AVAILABLE and hasattr(self, 'icons'):
            refresh_btn = ttk.Button(button_frame, text="Refresh Scripts", image=self.icons.get('scripts_refresh'), 
                                    compound=tk.LEFT, command=self.refresh_script_list)
        else:
            refresh_btn = ttk.Button(button_frame, text="Refresh Scripts", command=self.refresh_script_list)
        refresh_btn.pack(side='right')
        
        # Create a LabelFrame for favorites section
        favorites_frame = ttk.LabelFrame(list_frame, text="Favorites")
        favorites_frame.pack(side='top', fill='x', padx=5, pady=(5,0))
        
        columns = ('Favorite', 'Script Name')
        self.favorites_tree = ttk.Treeview(favorites_frame, columns=columns, show='headings', height=5)
        self.favorites_tree.pack(fill='x', padx=5, pady=5, expand=True)
        
        # Set favorites columns
        self.favorites_tree.heading('Favorite', text='♡')
        self.favorites_tree.heading('Script Name', text='Script Name')
        self.favorites_tree.column('Favorite', width=30, anchor='center', stretch=False)
        self.favorites_tree.column('Script Name', width=400, stretch=True)
        self.favorites_tree.bind('<Button-1>', self.on_tree_click)
        self.favorites_tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.favorites_tree.bind('<Button-3>', self.on_script_right_click)
        
        # Add spacing between frames
        ttk.Frame(list_frame, height=10).pack(fill='x')
        
        # Create a LabelFrame for all scripts section
        all_scripts_frame = ttk.LabelFrame(list_frame, text="All Scripts")
        all_scripts_frame.pack(side='top', fill='both', expand=True, padx=5, pady=(5,0))
        
        self.scripts_tree = ttk.Treeview(all_scripts_frame, columns=columns, show='headings')
        self.scripts_tree.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Set scripts columns
        self.scripts_tree.heading('Favorite', text='♡')
        self.scripts_tree.heading('Script Name', text='Script Name')
        self.scripts_tree.column('Favorite', width=30, anchor='center', stretch=False)
        self.scripts_tree.column('Script Name', width=400, stretch=True)
        self.scripts_tree.bind('<Button-1>', self.on_tree_click)
        self.scripts_tree.bind('<<TreeviewSelect>>', self.on_tree_select)
        self.scripts_tree.bind('<Button-3>', self.on_script_right_click)
        
        # Create context menu for scripts
        self.script_context_menu = tk.Menu(self.root, tearoff=0)
        self.script_context_menu.add_command(label="Toggle Favorite", command=self.toggle_script_favorite)
        self.script_context_menu.add_separator()
        self.script_context_menu.add_command(label="Run Script", command=self.run_script)
        self.script_context_menu.add_command(label="Run As...", command=self.run_script_as)
        self.script_context_menu.add_separator()
        self.script_context_menu.add_command(label="Open in Notepad", command=self.open_in_notepad)
        
    def setup_powershell_tab(self):
        # Create main frame for PowerShell information
        main_frame = ttk.Frame(self.powershell_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create a LabelFrame for PowerShell versions
        versions_frame = ttk.LabelFrame(main_frame, text="PowerShell Versions")
        versions_frame.pack(fill='x', expand=False, pady=(0, 10))
        
        # Function to get PowerShell versions and details
        def get_powershell_details():
            details = {}
            
            # Initialize installed PowerShell variants list
            details['installed_variants'] = []
            
            # Get Windows PowerShell version
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', '(Get-Host).Version.ToString()'],
                    capture_output=True,
                    text=True
                )
                win_ps_version = result.stdout.strip()
                details['win_ps_version'] = win_ps_version
                
                # Add to installed variants
                details['installed_variants'].append({
                    'name': 'Windows PowerShell',
                    'version': win_ps_version,
                    'path': 'powershell.exe',
                    'icon': 'powershell.exe'
                })
                
                # Check if Windows PowerShell update is available (PowerShell 5.1 is the latest for Windows PowerShell)
                # This checks if current version is less than 5.1
                current_version = [int(x) for x in details['win_ps_version'].split('.')]
                if len(current_version) >= 2 and (current_version[0] < 5 or (current_version[0] == 5 and current_version[1] < 1)):
                    details['win_ps_update'] = True
                else:
                    details['win_ps_update'] = False
            except:
                details['win_ps_version'] = "Not detected"
                details['win_ps_update'] = False
                
            # Check if PowerShell ISE is installed
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', "if (Get-Command powershell_ise.exe -ErrorAction SilentlyContinue) { (Get-Host).Version.ToString() } else { 'Not installed' }"],
                    capture_output=True,
                    text=True
                )
                ise_version = result.stdout.strip()
                details['ise_version'] = ise_version
                
                # Add to installed variants if ISE is installed
                if ise_version != 'Not installed':
                    details['installed_variants'].append({
                        'name': 'PowerShell ISE',
                        'version': ise_version,
                        'path': 'powershell_ise.exe',
                        'icon': 'powershell_ise.exe'
                    })
            except:
                details['ise_version'] = "Not detected"
                
            # Check if PowerShell Core (pwsh) is installed
            try:
                result = subprocess.run(
                    ['pwsh', '-Command', '(Get-Host).Version.ToString()'],
                    capture_output=True,
                    text=True
                )
                core_ps_version = result.stdout.strip()
                details['core_ps_version'] = core_ps_version
                
                # Add to installed variants
                details['installed_variants'].append({
                    'name': 'PowerShell Core',
                    'version': core_ps_version,
                    'path': 'pwsh.exe',
                    'icon': 'pwsh.exe'
                })
                
                # Check if PowerShell Core update is available
                # This requires an internet connection to check the GitHub API
                try:
                    # Check latest version from GitHub API
                    result = subprocess.run(
                        ['powershell.exe', '-Command', 
                         "try { $releaseInfo = Invoke-RestMethod -Uri 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -TimeoutSec 3; $releaseInfo.tag_name.TrimStart('v') } catch { 'Unknown' }"],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    latest_version = result.stdout.strip()
                    
                    if latest_version != 'Unknown':
                        current_parts = [int(x) for x in details['core_ps_version'].split('.')]
                        latest_parts = [int(x) for x in latest_version.split('.')]
                        
                        # Compare versions
                        update_available = False
                        for i in range(min(len(current_parts), len(latest_parts))):
                            if latest_parts[i] > current_parts[i]:
                                update_available = True
                                break
                            elif current_parts[i] > latest_parts[i]:
                                break
                        
                        details['core_ps_update'] = update_available
                        details['core_ps_latest'] = latest_version
                    else:
                        details['core_ps_update'] = False
                        details['core_ps_latest'] = "Unknown"
                except:
                    details['core_ps_update'] = False
                    details['core_ps_latest'] = "Unknown"
            except:
                details['core_ps_version'] = "Not installed"
                details['core_ps_update'] = False
                details['core_ps_latest'] = "N/A"
                
            # Check for other PowerShell preview versions
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', 
                     "if (Get-Command pwsh-preview -ErrorAction SilentlyContinue) { pwsh-preview -Command '(Get-Host).Version.ToString()' } else { 'Not installed' }"],
                    capture_output=True,
                    text=True
                )
                preview_version = result.stdout.strip()
                
                if preview_version != 'Not installed':
                    # Add to installed variants
                    details['installed_variants'].append({
                        'name': 'PowerShell Preview',
                        'version': preview_version,
                        'path': 'pwsh-preview.exe',
                        'icon': 'pwsh-preview.exe'
                    })
            except:
                pass
                
            # Get execution policy
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', 'Get-ExecutionPolicy'],
                    capture_output=True,
                    text=True
                )
                details['execution_policy'] = result.stdout.strip()
            except:
                details['execution_policy'] = "Unknown"
                
            # Get module path
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', '$env:PSModulePath'],
                    capture_output=True,
                    text=True
                )
                details['module_path'] = result.stdout.strip()
            except:
                details['module_path'] = "Unknown"
                
            # Get profile path
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', '$PROFILE'],
                    capture_output=True,
                    text=True
                )
                details['profile_path'] = result.stdout.strip()
            except:
                details['profile_path'] = "Unknown"
                
            # Get PSEdition
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', '$PSVersionTable.PSEdition'],
                    capture_output=True,
                    text=True
                )
                details['ps_edition'] = result.stdout.strip()
            except:
                details['ps_edition'] = "Unknown"
                
            # Get PS Platform
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', '$PSVersionTable.Platform'],
                    capture_output=True,
                    text=True
                )
                details['ps_platform'] = result.stdout.strip()
            except:
                details['ps_platform'] = "Unknown"
                
            return details
        
        # Get PowerShell details
        ps_details = get_powershell_details()
        
        # Create version info display
        version_info = ttk.Frame(versions_frame)
        version_info.pack(fill='x', padx=10, pady=10)
        
        # Create a subframe with scrollable area for installed PowerShell variants
        variants_container = ttk.Frame(version_info)
        variants_container.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=5)
        
        # Display all installed PowerShell variants
        row_num = 1
        for variant in ps_details['installed_variants']:
            # Variant name and version
            ttk.Label(variants_container, text=f"{variant['name']}:").grid(
                row=row_num, column=0, sticky='w', padx=(20, 10), pady=2)
            ttk.Label(variants_container, text=variant['version']).grid(
                row=row_num, column=1, sticky='w', pady=2)
            
            # Add launch button for this PowerShell variant
            def create_launch_command(path):
                return lambda: subprocess.Popen([path])
            
            launch_btn = ttk.Button(variants_container, text="Launch", 
                                  command=create_launch_command(variant['path']))
            launch_btn.grid(row=row_num, column=2, sticky='e', padx=5, pady=2)
            
            # Get status from the background check if available
            status_info = self.powershell_status.get(variant['name'], {'status': 'checking', 'version': ''})
            
            # Display status based on background check
            if status_info['status'] == 'checking':
                status_label = tk.Label(variants_container, text="Checking...", foreground="blue")
                status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
            elif status_info['status'] == 'update_available':
                version_text = f" ({status_info['version']})" if status_info['version'] else ""
                status_label = tk.Label(variants_container, text=f"Update Available{version_text}", foreground="green")
                status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
                # Add update button based on which PowerShell variant
                if variant['name'] == 'Windows PowerShell' or variant['name'] == 'PowerShell ISE':
                    def update_windows_powershell():
                        if hasattr(ctypes, 'windll'):
                            ctypes.windll.shell32.ShellExecuteW(
                                None, 
                                "runas",
                                "powershell.exe",
                                "-Command Start-Process -Wait powershell -ArgumentList '-Command Install-Module -Name PowerShellGet -Force -AllowClobber'",
                                None, 
                                1
                            )
                    update_btn = ttk.Button(variants_container, text="Update", command=update_windows_powershell)
                    update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                    
                elif variant['name'] == 'PowerShell Core':
                    def update_powershell_core():
                        if hasattr(ctypes, 'windll'):
                            ctypes.windll.shell32.ShellExecuteW(
                                None, 
                                "runas",
                                "powershell.exe",
                                "-Command Start-Process -Wait msedge -ArgumentList 'https://github.com/PowerShell/PowerShell/releases/latest'",
                                None, 
                                1
                            )
                    update_btn = ttk.Button(variants_container, text="Update", command=update_powershell_core)
                    update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                    
            elif status_info['status'] == 'up_to_date':
                status_label = tk.Label(variants_container, text="Up to Date", foreground="green")
                status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
            elif status_info['status'] == 'preview':
                status_label = tk.Label(variants_container, text="Preview Version", foreground="purple")
                status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
            elif status_info['status'] == 'unknown':
                status_label = tk.Label(variants_container, text="Status Unknown", foreground="orange")
                status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
            # Add legacy check for compatibility with previous code
            # This will be used until the background check completes
            elif variant['name'] == 'Windows PowerShell' and ps_details['win_ps_update']:
                update_label = tk.Label(variants_container, text="Update Available", foreground="green")
                update_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
                def update_windows_powershell():
                    if hasattr(ctypes, 'windll'):
                        ctypes.windll.shell32.ShellExecuteW(
                            None, 
                            "runas",
                            "powershell.exe",
                            "-Command Start-Process -Wait powershell -ArgumentList '-Command Install-Module -Name PowerShellGet -Force -AllowClobber'",
                            None, 
                            1
                        )
                update_btn = ttk.Button(variants_container, text="Update", command=update_windows_powershell)
                update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                
            elif variant['name'] == 'PowerShell Core' and ps_details['core_ps_update']:
                update_label = tk.Label(variants_container, text=f"Update Available ({ps_details['core_ps_latest']})", foreground="green")
                update_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                
                def update_powershell_core():
                    if hasattr(ctypes, 'windll'):
                        ctypes.windll.shell32.ShellExecuteW(
                            None, 
                            "runas",
                            "powershell.exe",
                            "-Command Start-Process -Wait msedge -ArgumentList 'https://github.com/PowerShell/PowerShell/releases/latest'",
                            None, 
                            1
                        )
                update_btn = ttk.Button(variants_container, text="Update", command=update_powershell_core)
                update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
            
            row_num += 1
        
            # Add some spacing at the bottom for better visual appearance
            ttk.Frame(version_info, height=5).grid(row=row_num + 1, column=0, columnspan=5, sticky='ew')        # Create a LabelFrame for execution policy
        policy_frame = ttk.LabelFrame(main_frame, text="Execution Policy")
        policy_frame.pack(fill='x', expand=False, pady=(0, 10))
        
        # Create policy info display
        policy_info = ttk.Frame(policy_frame)
        policy_info.pack(fill='x', padx=10, pady=10)
        
        # Current Policy
        ttk.Label(policy_info, text="Current Policy:").grid(row=0, column=0, sticky='w', padx=(0, 10), pady=2)
        
        # Use different colors based on policy type
        policy_value = ps_details['execution_policy']
        policy_color = "green" if policy_value.lower() in ['remotesigned', 'unrestricted', 'bypass'] else "red"
        
        policy_label = tk.Label(policy_info, text=policy_value, foreground=policy_color)
        policy_label.grid(row=0, column=1, sticky='w', pady=2)
        
        # Policy description
        policy_desc = {
            "Restricted": "Does not load configuration files or run scripts. Default state.",
            "AllSigned": "Scripts can run, but requires all scripts and config files to be signed by a trusted publisher.",
            "RemoteSigned": "Scripts downloaded from the internet must be signed by a trusted publisher.",
            "Unrestricted": "Unsigned scripts can run (a warning appears before running scripts from the internet).",
            "Bypass": "Nothing is blocked and there are no warnings or prompts.",
            "Undefined": "No execution policy set in the current scope.",
            "Default": "Sets the default execution policy (Restricted for Windows clients)."
        }
        
        desc = policy_desc.get(policy_value, "Unknown policy")
        ttk.Label(policy_info, text="Description:").grid(row=1, column=0, sticky='nw', padx=(0, 10), pady=2)
        desc_label = ttk.Label(policy_info, text=desc, wraplength=400)
        desc_label.grid(row=1, column=1, sticky='w', pady=2)
        
        # Add button to open powershell as admin to change policy
        def open_ps_admin():
            if hasattr(ctypes, 'windll'):
                ctypes.windll.shell32.ShellExecuteW(
                    None, 
                    "runas",
                    "powershell.exe",
                    "-Command Start-Process powershell -Verb RunAs",
                    None, 
                    1
                )
        
        # Add a button frame to contain both buttons side by side
        button_frame = ttk.Frame(policy_info)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        # Add button to open PowerShell as Admin
        ttk.Button(button_frame, text="Open PowerShell as Admin", command=open_ps_admin).pack(side='left', padx=(0, 5))
        
        # Add button to set execution policy for current user
        def set_current_user_policy():
            policy = "RemoteSigned"  # Most common safe policy
            result = messagebox.askyesno(
                "Change Execution Policy", 
                f"Do you want to change the execution policy for the current user to '{policy}'?\n\n" +
                "This will allow scripts to run without requiring administrator privileges."
            )
            
            if result:
                try:
                    # Run the command to change execution policy for current user
                    subprocess.run(
                        ['powershell.exe', '-Command', f"Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy {policy} -Force"],
                        capture_output=True,
                        text=True
                    )
                    messagebox.showinfo(
                        "Success", 
                        f"Execution policy for current user has been set to '{policy}'.\n\n" +
                        "Please refresh the PowerShell info to see the changes."
                    )
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to change execution policy:\n{str(e)}")
        
        ttk.Button(button_frame, text="Set for Current User", command=set_current_user_policy).pack(side='left')
        
        # Create a LabelFrame for PowerShell paths
        paths_frame = ttk.LabelFrame(main_frame, text="PowerShell Paths")
        paths_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Create paths info display with scrollable text widget
        paths_info = ttk.Frame(paths_frame)
        paths_info.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Profile Path
        ttk.Label(paths_info, text="Profile Path:").grid(row=0, column=0, sticky='w', padx=(0, 10), pady=2)
        profile_text = ttk.Entry(paths_info, width=50)
        profile_text.insert(0, ps_details['profile_path'])
        profile_text.configure(state='readonly')
        profile_text.grid(row=0, column=1, sticky='w', pady=2)
        
        # Module Path
        ttk.Label(paths_info, text="Module Path:").grid(row=1, column=0, sticky='nw', padx=(0, 10), pady=2)
        
        # Create scrollable text widget for module paths
        module_text = tk.Text(paths_info, height=5, width=50, wrap='none')
        module_text.insert('1.0', ps_details['module_path'].replace(';', ';\n'))
        module_text.configure(state='disabled')
        module_text.grid(row=1, column=1, sticky='w', pady=2)
        
        # Add scrollbar
        module_scroll = ttk.Scrollbar(paths_info, orient='vertical', command=module_text.yview)
        module_scroll.grid(row=1, column=2, sticky='ns')
        module_text.configure(yscrollcommand=module_scroll.set)
        
        # Button to refresh PowerShell info
        def refresh_powershell_info():
            # Get updated PowerShell details
            updated_ps_details = get_powershell_details()
            
            # Clear version info and recreate
            for widget in version_info.winfo_children():
                widget.destroy()
            
            # Create a subframe with scrollable area for installed PowerShell variants
            variants_container = ttk.Frame(version_info)
            variants_container.grid(row=0, column=0, columnspan=4, sticky='nsew', pady=5)
            
            # Header for installed variants
            ttk.Label(variants_container, text="Installed PowerShell Variants:", font=('TkDefaultFont', 10, 'bold')).grid(
                row=0, column=0, columnspan=4, sticky='w', pady=(0, 5))
            
            # Display all installed PowerShell variants
            row_num = 1
            for variant in updated_ps_details['installed_variants']:
                # Variant name and version
                ttk.Label(variants_container, text=f"{variant['name']}:").grid(
                    row=row_num, column=0, sticky='w', padx=(20, 10), pady=2)
                ttk.Label(variants_container, text=variant['version']).grid(
                    row=row_num, column=1, sticky='w', pady=2)
                
                # Add launch button for this PowerShell variant
                def create_launch_command(path):
                    return lambda: subprocess.Popen([path])
                
                launch_btn = ttk.Button(variants_container, text="Launch", 
                                      command=create_launch_command(variant['path']))
                launch_btn.grid(row=row_num, column=2, sticky='e', padx=5, pady=2)
                
                # Get status from the background check if available
                status_info = self.powershell_status.get(variant['name'], {'status': 'checking', 'version': ''})
                
                # Display status based on background check
                if status_info['status'] == 'checking':
                    status_label = tk.Label(variants_container, text="Checking...", foreground="blue")
                    status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                elif status_info['status'] == 'update_available':
                    version_text = f" ({status_info['version']})" if status_info['version'] else ""
                    status_label = tk.Label(variants_container, text=f"Update Available{version_text}", foreground="green")
                    status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                    # Add update button based on which PowerShell variant
                    if variant['name'] == 'Windows PowerShell' or variant['name'] == 'PowerShell ISE':
                        def update_windows_powershell():
                            if hasattr(ctypes, 'windll'):
                                ctypes.windll.shell32.ShellExecuteW(
                                    None, 
                                    "runas",
                                    "powershell.exe",
                                    "-Command Start-Process -Wait powershell -ArgumentList '-Command Install-Module -Name PowerShellGet -Force -AllowClobber'",
                                    None, 
                                    1
                                )
                        update_btn = ttk.Button(variants_container, text="Update", command=update_windows_powershell)
                        update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                        
                    elif variant['name'] == 'PowerShell Core':
                        def update_powershell_core():
                            if hasattr(ctypes, 'windll'):
                                ctypes.windll.shell32.ShellExecuteW(
                                    None, 
                                    "runas",
                                    "powershell.exe",
                                    "-Command Start-Process -Wait msedge -ArgumentList 'https://github.com/PowerShell/PowerShell/releases/latest'",
                                    None, 
                                    1
                                )
                        update_btn = ttk.Button(variants_container, text="Update", command=update_powershell_core)
                        update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                        
                elif status_info['status'] == 'up_to_date':
                    status_label = tk.Label(variants_container, text="Up to Date", foreground="green")
                    status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                elif status_info['status'] == 'preview':
                    status_label = tk.Label(variants_container, text="Preview Version", foreground="purple")
                    status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                elif status_info['status'] == 'unknown':
                    status_label = tk.Label(variants_container, text="Status Unknown", foreground="orange")
                    status_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                # Add legacy check for compatibility with previous code
                # This will be used until the background check completes
                elif variant['name'] == 'Windows PowerShell' and updated_ps_details['win_ps_update']:
                    update_label = tk.Label(variants_container, text="Update Available", foreground="green")
                    update_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                    def update_windows_powershell():
                        if hasattr(ctypes, 'windll'):
                            ctypes.windll.shell32.ShellExecuteW(
                                None, 
                                "runas",
                                "powershell.exe",
                                "-Command Start-Process -Wait powershell -ArgumentList '-Command Install-Module -Name PowerShellGet -Force -AllowClobber'",
                                None, 
                                1
                            )
                    update_btn = ttk.Button(variants_container, text="Update", command=update_windows_powershell)
                    update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                    
                elif variant['name'] == 'PowerShell Core' and updated_ps_details['core_ps_update']:
                    update_label = tk.Label(variants_container, text=f"Update Available ({updated_ps_details['core_ps_latest']})", foreground="green")
                    update_label.grid(row=row_num, column=3, sticky='w', padx=5, pady=2)
                    
                    def update_powershell_core():
                        if hasattr(ctypes, 'windll'):
                            ctypes.windll.shell32.ShellExecuteW(
                                None, 
                                "runas",
                                "powershell.exe",
                                "-Command Start-Process -Wait msedge -ArgumentList 'https://github.com/PowerShell/PowerShell/releases/latest'",
                                None, 
                                1
                            )
                    update_btn = ttk.Button(variants_container, text="Update", command=update_powershell_core)
                    update_btn.grid(row=row_num, column=4, sticky='e', padx=5, pady=2)
                
                row_num += 1
            
            # Add some spacing at the bottom for better visual appearance
            ttk.Frame(version_info, height=5).grid(row=row_num + 1, column=0, columnspan=5, sticky='ew')
            
            # Clear policy info and recreate
            for widget in policy_info.winfo_children():
                widget.destroy()
            
            # Current Policy
            ttk.Label(policy_info, text="Current Policy:").grid(row=0, column=0, sticky='w', padx=(0, 10), pady=2)
            
            # Use different colors based on policy type
            policy_value = updated_ps_details['execution_policy']
            policy_color = "green" if policy_value.lower() in ['remotesigned', 'unrestricted', 'bypass'] else "red"
            
            policy_label = tk.Label(policy_info, text=policy_value, foreground=policy_color)
            policy_label.grid(row=0, column=1, sticky='w', pady=2)
            
            # Policy description
            policy_desc = {
                "Restricted": "Does not load configuration files or run scripts. Default state.",
                "AllSigned": "Scripts can run, but requires all scripts and config files to be signed by a trusted publisher.",
                "RemoteSigned": "Scripts downloaded from the internet must be signed by a trusted publisher.",
                "Unrestricted": "Unsigned scripts can run (a warning appears before running scripts from the internet).",
                "Bypass": "Nothing is blocked and there are no warnings or prompts.",
                "Undefined": "No execution policy set in the current scope.",
                "Default": "Sets the default execution policy (Restricted for Windows clients)."
            }
            
            desc = policy_desc.get(policy_value, "Unknown policy")
            ttk.Label(policy_info, text="Description:").grid(row=1, column=0, sticky='nw', padx=(0, 10), pady=2)
            desc_label = ttk.Label(policy_info, text=desc, wraplength=400)
            desc_label.grid(row=1, column=1, sticky='w', pady=2)
            
            # Add a button frame to contain both buttons side by side
            button_frame = ttk.Frame(policy_info)
            button_frame.grid(row=2, column=0, columnspan=2, pady=10)
            
            # Add button to open PowerShell as Admin
            ttk.Button(button_frame, text="Open PowerShell as Admin", command=open_ps_admin).pack(side='left', padx=(0, 5))
            
            # Add button to set execution policy for current user
            ttk.Button(button_frame, text="Set for Current User", command=set_current_user_policy).pack(side='left')
            
            # Clear paths info and recreate
            for widget in paths_info.winfo_children():
                widget.destroy()
            
            # Profile Path
            ttk.Label(paths_info, text="Profile Path:").grid(row=0, column=0, sticky='w', padx=(0, 10), pady=2)
            profile_text = ttk.Entry(paths_info, width=50)
            profile_text.insert(0, updated_ps_details['profile_path'])
            profile_text.configure(state='readonly')
            profile_text.grid(row=0, column=1, sticky='w', pady=2)
            
            # Module Path
            ttk.Label(paths_info, text="Module Path:").grid(row=1, column=0, sticky='nw', padx=(0, 10), pady=2)
            
            # Create scrollable text widget for module paths
            module_text = tk.Text(paths_info, height=5, width=50, wrap='none')
            module_text.insert('1.0', updated_ps_details['module_path'].replace(';', ';\n'))
            module_text.configure(state='disabled')
            module_text.grid(row=1, column=1, sticky='w', pady=2)
            
            # Add scrollbar
            module_scroll = ttk.Scrollbar(paths_info, orient='vertical', command=module_text.yview)
            module_scroll.grid(row=1, column=2, sticky='ns')
            module_text.configure(yscrollcommand=module_scroll.set)
            
            # Show notification about refresh
            toast = Notification(
                app_id="PowerShell Script Manager",
                title="PowerShell Info Refreshed",
                msg="PowerShell information has been refreshed",
                duration="short"
            )
            toast.set_audio(audio.Default, loop=False)
            toast.show()
        
        refresh_btn = ttk.Button(main_frame, text="Refresh PowerShell Info", command=refresh_powershell_info)
        refresh_btn.pack(pady=(0, 10))
        
    def setup_modules_tab(self):
        """Set up the PowerShell Modules tab to display installed modules and their details"""
        # Create main frame for modules information
        main_frame = ttk.Frame(self.modules_tab)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create top section with filter and buttons
        top_frame = ttk.Frame(main_frame)
        top_frame.pack(fill='x', pady=(0, 10))
        
        # Filter controls
        ttk.Label(top_frame, text="Filter:").pack(side='left', padx=(0, 5))
        self.module_filter_var = tk.StringVar()
        filter_entry = ttk.Entry(top_frame, textvariable=self.module_filter_var, width=30)
        filter_entry.pack(side='left', padx=(0, 5))
        
        # Create function to filter modules when the entry changes
        def on_filter_change(*args):
            self.filter_modules()
        
        self.module_filter_var.trace_add('write', on_filter_change)
        
        # Create a refresh button
        def refresh_modules():
            self.load_modules()
            # Show notification about refresh
            toast = Notification(
                app_id="PowerShell Script Manager",
                title="Modules Refreshed",
                msg="PowerShell modules have been refreshed",
                duration="short"
            )
            toast.set_audio(audio.Default, loop=False)
            toast.show()
        
        refresh_btn = ttk.Button(top_frame, text="Refresh Modules", command=refresh_modules)
        refresh_btn.pack(side='right')
        
        # Create a LabelFrame for the modules list
        modules_frame = ttk.LabelFrame(main_frame, text="Installed PowerShell Modules")
        modules_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Create treeview for modules list
        columns = ('name', 'version', 'description', 'path', 'repository')
        self.modules_tree = ttk.Treeview(modules_frame, columns=columns, show='headings', selectmode='browse')
        self.modules_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        
        # Configure columns
        self.modules_tree.heading('name', text='Name', command=lambda: self.treeview_sort_column(self.modules_tree, 'name', False))
        self.modules_tree.heading('version', text='Version', command=lambda: self.treeview_sort_column(self.modules_tree, 'version', False))
        self.modules_tree.heading('description', text='Description', command=lambda: self.treeview_sort_column(self.modules_tree, 'description', False))
        self.modules_tree.heading('path', text='Path', command=lambda: self.treeview_sort_column(self.modules_tree, 'path', False))
        self.modules_tree.heading('repository', text='Repository', command=lambda: self.treeview_sort_column(self.modules_tree, 'repository', False))
        
        # Set column widths
        self.modules_tree.column('name', width=150, stretch=False)
        self.modules_tree.column('version', width=80, stretch=False)
        self.modules_tree.column('description', width=300, stretch=True)
        self.modules_tree.column('path', width=200, stretch=True)
        self.modules_tree.column('repository', width=120, stretch=False)
        
        # Add vertical scrollbar
        y_scrollbar = ttk.Scrollbar(modules_frame, orient='vertical', command=self.modules_tree.yview)
        y_scrollbar.pack(side='right', fill='y')
        self.modules_tree.configure(yscrollcommand=y_scrollbar.set)
        
        # Add horizontal scrollbar
        x_scrollbar = ttk.Scrollbar(main_frame, orient='horizontal', command=self.modules_tree.xview)
        x_scrollbar.pack(side='bottom', fill='x')
        self.modules_tree.configure(xscrollcommand=x_scrollbar.set)
        
        # Bind double-click to show module details
        self.modules_tree.bind('<Double-1>', self.show_module_details)
        
        # Status display to show loading information
        self.modules_status_var = tk.StringVar(value="Loading modules...")
        status_label = ttk.Label(main_frame, textvariable=self.modules_status_var)
        status_label.pack(side='left', pady=(5, 0))
        
        # Create a context menu for modules
        self.module_context_menu = tk.Menu(self.modules_tree, tearoff=0)
        self.module_context_menu.add_command(label="Show Details", command=self.show_module_details_context)
        self.module_context_menu.add_command(label="Update Module", command=self.update_module)
        self.module_context_menu.add_command(label="Uninstall Module", command=self.uninstall_module)
        self.module_context_menu.add_separator()
        self.module_context_menu.add_command(label="Copy Name", command=lambda: self.copy_module_info('name'))
        self.module_context_menu.add_command(label="Copy Version", command=lambda: self.copy_module_info('version'))
        self.module_context_menu.add_command(label="Copy Path", command=lambda: self.copy_module_info('path'))
        
        # Bind right-click to show context menu
        self.modules_tree.bind('<Button-3>', self.show_module_context_menu)
        
        # Load modules on tab creation
        self.load_modules()
    
    def treeview_sort_column(self, tree, col, reverse):
        """Sort treeview contents when a column header is clicked"""
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            # Try to sort numerically for version columns
            if col == 'version':
                # Custom version sorting that understands semantic versioning
                def version_key(version_str):
                    # Extract components from version string (e.g., "1.2.3-alpha.1")
                    parts = version_str[0].split('-')
                    version_nums = parts[0].split('.')
                    # Convert numbers to integers for proper sorting
                    nums = []
                    for v in version_nums:
                        try:
                            nums.append(int(v))
                        except ValueError:
                            nums.append(0)
                    # Pad with zeros to ensure consistent comparison length
                    while len(nums) < 4:
                        nums.append(0)
                    # Add pre-release tag with lower priority than release versions
                    prerelease_val = 0
                    if len(parts) > 1:
                        prerelease_val = -1  # Pre-release versions sort before release versions
                    return (nums[0], nums[1], nums[2], nums[3], prerelease_val)
                
                l.sort(key=version_key, reverse=reverse)
            else:
                # Normal alphabetical sorting for other columns
                l.sort(reverse=reverse)
        except Exception:
            # Fall back to default sorting if numeric sort fails
            l.sort(reverse=reverse)
            
        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)

        # Reverse sort next time
        tree.heading(col, command=lambda: self.treeview_sort_column(tree, col, not reverse))
    
    def show_module_context_menu(self, event):
        """Show context menu for a module item"""
        item = self.modules_tree.identify_row(event.y)
        if item:
            # Select the item that was right-clicked
            self.modules_tree.selection_set(item)
            self.module_context_menu.post(event.x_root, event.y_root)
    
    def copy_module_info(self, column):
        """Copy the selected module's information to clipboard"""
        selection = self.modules_tree.selection()
        if selection:
            item = selection[0]
            info = self.modules_tree.item(item, 'values')[{'name': 0, 'version': 1, 'path': 3}[column]]
            self.root.clipboard_clear()
            self.root.clipboard_append(info)
    
    def show_module_details_context(self, event=None):
        """Show details for the selected module (context menu version)"""
        selection = self.modules_tree.selection()
        if selection:
            self.show_module_details(item=selection[0])
    
    def show_module_details(self, event=None, item=None):
        """Show detailed information for the selected module"""
        if event:
            item = self.modules_tree.identify_row(event.y)
            
        if not item:
            selection = self.modules_tree.selection()
            if not selection:
                return
            item = selection[0]
            
        # Get module name from selection
        module_name = self.modules_tree.item(item, 'values')[0]
        
        # Create a new toplevel window for details
        details_window = tk.Toplevel(self.root)
        details_window.title(f"Module Details: {module_name}")
        details_window.geometry("700x500")
        details_window.minsize(600, 400)
        
        # Make the window modal
        details_window.transient(self.root)
        details_window.grab_set()
        
        # Create main frame
        main_frame = ttk.Frame(details_window, padding=10)
        main_frame.pack(fill='both', expand=True)
        
        # Start fetching detailed information about the module
        info_frame = ttk.LabelFrame(main_frame, text="Module Information")
        info_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Show loading indicator
        loading_label = ttk.Label(info_frame, text="Loading module details...")
        loading_label.pack(pady=20)
        
        # Function to get detailed module information
        def get_module_details():
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', f"Get-Module -Name '{module_name}' -ListAvailable | Select-Object Name, Version, Description, Path, Author, CompanyName, Copyright, PowerShellVersion, CompatiblePSEditions, PrivateData | ConvertTo-Json"],
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                
                # Parse the JSON output
                module_info = json.loads(result.stdout)
                
                # If multiple modules were returned, use the first one
                if isinstance(module_info, list):
                    module_info = module_info[0]
                
                # Get exported commands (functions, cmdlets, aliases)
                commands_result = subprocess.run(
                    ['powershell.exe', '-Command', f"Get-Command -Module '{module_name}' | Select-Object Name, CommandType, Version | ConvertTo-Json"],
                    capture_output=True,
                    text=True,
                    timeout=10
                )
                
                try:
                    commands_info = json.loads(commands_result.stdout)
                    if not isinstance(commands_info, list):
                        commands_info = [commands_info]
                except:
                    commands_info = []
                
                return module_info, commands_info
            except Exception as e:
                return {"Error": str(e)}, []
        
        # Function to update the UI with module details
        def update_module_details_ui():
            # Remove loading label
            loading_label.destroy()
            
            # Get module details in a separate thread
            module_info, commands_info = get_module_details()
            
            # Create a notebook for details tabs
            details_notebook = ttk.Notebook(info_frame)
            details_notebook.pack(fill='both', expand=True, padx=5, pady=5)
            
            # General tab
            general_tab = ttk.Frame(details_notebook)
            details_notebook.add(general_tab, text="General")
            
            # Properties tab
            props_tab = ttk.Frame(details_notebook)
            details_notebook.add(props_tab, text="Properties")
            
            # Commands tab
            commands_tab = ttk.Frame(details_notebook)
            details_notebook.add(commands_tab, text="Commands")
            
            # Fill general tab with basic information
            row = 0
            for label, key in [
                ("Name:", "Name"),
                ("Version:", "Version"),
                ("Description:", "Description"),
                ("Author:", "Author"),
                ("Company:", "CompanyName"),
                ("Copyright:", "Copyright"),
                ("Path:", "Path"),
            ]:
                ttk.Label(general_tab, text=label, font=('TkDefaultFont', 10, 'bold')).grid(
                    row=row, column=0, sticky='nw', padx=(10, 5), pady=5)
                value = module_info.get(key, "N/A")
                if key == "Description" and value != "N/A":
                    desc_text = tk.Text(general_tab, wrap='word', height=4, width=50)
                    desc_text.insert('1.0', value)
                    desc_text.configure(state='disabled')
                    desc_text.grid(row=row, column=1, sticky='nw', padx=5, pady=5)
                else:
                    ttk.Label(general_tab, text=value, wraplength=400).grid(
                        row=row, column=1, sticky='nw', padx=5, pady=5)
                row += 1
            
            # Fill properties tab with additional information
            row = 0
            for label, key in [
                ("PowerShell Version:", "PowerShellVersion"),
                ("Compatible PS Editions:", "CompatiblePSEditions"),
                ("Tags:", "PrivateData.PSData.Tags"),
                ("License URI:", "PrivateData.PSData.LicenseUri"),
                ("Project URI:", "PrivateData.PSData.ProjectUri"),
                ("Release Notes:", "PrivateData.PSData.ReleaseNotes"),
            ]:
                ttk.Label(props_tab, text=label, font=('TkDefaultFont', 10, 'bold')).grid(
                    row=row, column=0, sticky='nw', padx=(10, 5), pady=5)
                
                # Handle nested properties
                if "." in key:
                    parts = key.split(".")
                    value = module_info
                    try:
                        for part in parts:
                            if isinstance(value, dict) and part in value:
                                value = value[part]
                            else:
                                value = "N/A"
                                break
                    except:
                        value = "N/A"
                else:
                    value = module_info.get(key, "N/A")
                
                # Format list values
                if isinstance(value, list):
                    value = ", ".join(str(item) for item in value)
                
                # Format special values
                if key == "PrivateData.PSData.ProjectUri" and value != "N/A":
                    link_frame = ttk.Frame(props_tab)
                    link_frame.grid(row=row, column=1, sticky='nw', padx=5, pady=5)
                    
                    link_text = ttk.Label(link_frame, text=value, foreground="blue", cursor="hand2")
                    link_text.pack(side='left')
                    link_text.bind("<Button-1>", lambda e, url=value: webbrowser.open(url))
                else:
                    ttk.Label(props_tab, text=str(value), wraplength=400).grid(
                        row=row, column=1, sticky='nw', padx=5, pady=5)
                row += 1
            
            # Fill commands tab with command list
            if commands_info:
                # Create a treeview for commands
                cmd_columns = ('name', 'type', 'version')
                cmd_tree = ttk.Treeview(commands_tab, columns=cmd_columns, show='headings')
                cmd_tree.pack(side='left', fill='both', expand=True, padx=5, pady=5)
                
                # Add scrollbar
                cmd_scroll = ttk.Scrollbar(commands_tab, orient='vertical', command=cmd_tree.yview)
                cmd_scroll.pack(side='right', fill='y')
                cmd_tree.configure(yscrollcommand=cmd_scroll.set)
                
                # Configure columns
                cmd_tree.heading('name', text='Command Name')
                cmd_tree.heading('type', text='Type')
                cmd_tree.heading('version', text='Version')
                
                cmd_tree.column('name', width=250)
                cmd_tree.column('type', width=100)
                cmd_tree.column('version', width=80)
                
                # Add commands to the tree
                for cmd in commands_info:
                    cmd_tree.insert('', 'end', values=(
                        cmd.get('Name', 'N/A'),
                        cmd.get('CommandType', 'N/A'),
                        cmd.get('Version', 'N/A')
                    ))
            else:
                ttk.Label(commands_tab, text="No commands found or unable to retrieve command information").pack(pady=20)
        
        # Start a new thread to fetch module details
        threading.Thread(target=update_module_details_ui, daemon=True).start()
        
        # Close button
        close_button = ttk.Button(main_frame, text="Close", command=details_window.destroy)
        close_button.pack(side='right', pady=(5, 0))
        
        # Center the window on the screen
        details_window.update_idletasks()
        width = details_window.winfo_width()
        height = details_window.winfo_height()
        x = (details_window.winfo_screenwidth() // 2) - (width // 2)
        y = (details_window.winfo_screenheight() // 2) - (height // 2)
        details_window.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    
    def update_module(self):
        """Update the selected PowerShell module"""
        selection = self.modules_tree.selection()
        if not selection:
            return
            
        module_name = self.modules_tree.item(selection[0], 'values')[0]
        
        # Confirm update
        if not messagebox.askyesno("Update Module", f"Are you sure you want to update the module '{module_name}'?"):
            return
            
        # Run PowerShell command to update the module with elevated privileges
        if hasattr(ctypes, 'windll'):
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas",
                "powershell.exe",
                f"-Command Start-Process -Wait powershell -ArgumentList '-Command Update-Module -Name {module_name} -Force'",
                None, 
                1
            )
            
            # Refresh the modules list after update
            self.load_modules()
    
    def uninstall_module(self):
        """Uninstall the selected PowerShell module"""
        selection = self.modules_tree.selection()
        if not selection:
            return
            
        module_name = self.modules_tree.item(selection[0], 'values')[0]
        
        # Confirm uninstall
        if not messagebox.askyesno("Uninstall Module", 
                                  f"Are you sure you want to uninstall the module '{module_name}'?\n\n" +
                                  "This action cannot be undone!",
                                  icon='warning'):
            return
            
        # Run PowerShell command to uninstall the module with elevated privileges
        if hasattr(ctypes, 'windll'):
            ctypes.windll.shell32.ShellExecuteW(
                None, 
                "runas",
                "powershell.exe",
                f"-Command Start-Process -Wait powershell -ArgumentList '-Command Uninstall-Module -Name {module_name} -Force'",
                None, 
                1
            )
            
            # Refresh the modules list after uninstall
            self.load_modules()
    
    def load_modules(self):
        """Load PowerShell modules in a background thread"""
        # Clear the current modules list
        for item in self.modules_tree.get_children():
            self.modules_tree.delete(item)
        
        # Update status
        self.modules_status_var.set("Loading modules...")
        
        def get_modules_thread():
            try:
                # Get installed modules
                result = subprocess.run(
                    ['powershell.exe', '-Command', 
                     "Get-Module -ListAvailable | Select-Object Name, Version, Description, Path, RepositorySourceLocation | ConvertTo-Json -Depth 1"],
                    capture_output=True,
                    text=True,
                    timeout=30
                )
                
                try:
                    modules = json.loads(result.stdout)
                    # Ensure modules is a list even if only one module is returned
                    if not isinstance(modules, list):
                        modules = [modules]
                except json.JSONDecodeError:
                    modules = []
                
                # Update UI in the main thread
                self.root.after(0, lambda: self.update_modules_ui(modules))
                
            except Exception as e:
                # Update status with error
                self.root.after(0, lambda: self.modules_status_var.set(f"Error loading modules: {str(e)}"))
        
        # Start a thread to fetch modules
        threading.Thread(target=get_modules_thread, daemon=True).start()
    
    def update_modules_ui(self, modules):
        """Update the modules UI with the loaded modules"""
        # Clear existing items
        for item in self.modules_tree.get_children():
            self.modules_tree.delete(item)
        
        # Add modules to the treeview
        for module in modules:
            # Get repository source if available
            repo = "N/A"
            if isinstance(module.get('RepositorySourceLocation'), str):
                repo = module.get('RepositorySourceLocation')
            elif isinstance(module.get('RepositorySourceLocation'), dict) and 'Location' in module['RepositorySourceLocation']:
                repo = module['RepositorySourceLocation']['Location']
            
            self.modules_tree.insert('', 'end', values=(
                module.get('Name', 'N/A'),
                module.get('Version', 'N/A'),
                module.get('Description', 'N/A'),
                module.get('Path', 'N/A'),
                repo
            ))
        
        # Update status
        count = len(modules)
        self.modules_status_var.set(f"Total modules: {count}")
        
        # Apply any active filter
        self.filter_modules()
    
    def filter_modules(self):
        """Filter modules based on the search text"""
        search_term = self.module_filter_var.get().lower()
        
        # Show all items if search is empty
        if not search_term:
            for item in self.modules_tree.get_children():
                self.modules_tree.item(item, tags=())
                self.modules_tree.detach(item)
                self.modules_tree.reattach(item, '', 'end')
            return
        
        # Filter items
        visible_count = 0
        for item in self.modules_tree.get_children():
            values = self.modules_tree.item(item, 'values')
            if (search_term in values[0].lower() or  # Name
                search_term in values[1].lower() or  # Version
                search_term in values[2].lower()):   # Description
                self.modules_tree.item(item, tags=())
                self.modules_tree.detach(item)
                self.modules_tree.reattach(item, '', 'end')
                visible_count += 1
            else:
                self.modules_tree.detach(item)
        
        # Update status with filter info
        self.modules_status_var.set(f"Showing {visible_count} modules (filtered)")
        
    def setup_folders_tab(self):
        # Create main frame for settings
        settings_frame = ttk.Frame(self.folders_tab)
        settings_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Create frame for folder management
        folder_frame = ttk.LabelFrame(settings_frame, text="PowerShell Script Folders")
        folder_frame.pack(side='top', fill='both', expand=True, pady=(5, 5), padx=5)
        
        # Create button frame at the top
        button_frame = ttk.Frame(folder_frame)
        button_frame.pack(side='top', fill='x', pady=5)
        
        # Remove folder button with icon (first to appear rightmost)
        if PILLOW_AVAILABLE and hasattr(self, 'icons'):
            remove_btn = ttk.Button(button_frame, text="Remove Selected", image=self.icons.get('folder_remove'), 
                                   compound=tk.LEFT, command=self.remove_folder)
        else:
            remove_btn = ttk.Button(button_frame, text="Remove Selected", command=self.remove_folder)
        remove_btn.pack(side='right', padx=(5,10))
        
        # Add folder button with icon (second to appear, to the left of remove button)
        if PILLOW_AVAILABLE and hasattr(self, 'icons'):
            add_btn = ttk.Button(button_frame, text="Add Folder", image=self.icons.get('folder_add'), 
                                compound=tk.LEFT, command=self.add_folder)
        else:
            add_btn = ttk.Button(button_frame, text="Add Folder", command=self.add_folder)
        add_btn.pack(side='right')
        
        # Create listbox for folders
        self.folder_listbox = tk.Listbox(folder_frame, selectmode=tk.SINGLE)
        self.folder_listbox.pack(side='left', expand=True, fill='both', pady=(5,0))
        
        # Add right-click context menu to folder listbox
        self.folder_context_menu = tk.Menu(self.folder_listbox, tearoff=0)
        self.folder_context_menu.add_command(label="Remove Selected Folder", 
                                          command=self.remove_folder)
        self.folder_context_menu.add_command(label="Open Folder Location", 
                                          command=self.open_folder_location)
        self.folder_listbox.bind("<Button-3>", self.show_folder_context_menu)
        
        # Load saved folders
        self.refresh_folder_list()
        
    def add_folder(self):
        folder_path = filedialog.askdirectory(title="Select Folder")
        if folder_path:
            self.app_data.add_folder(folder_path)
            self.refresh_folder_list()
            self.refresh_script_list()
            
    def remove_folder(self):
        selection = self.folder_listbox.curselection()
        if selection:
            # Get the actual folder path from our stored list
            index = selection[0]
            if index < len(self.folder_paths):
                folder_path = self.folder_paths[index]
                
                # Count scripts in this folder before removing
                script_count = 0
                if os.path.exists(folder_path):
                    for root, _, files in os.walk(folder_path):
                        script_count += sum(1 for file in files if file.endswith('.ps1'))
                
                # Remove the folder
                self.app_data.remove_folder(folder_path)
                self.refresh_folder_list()
                # Refresh script list without showing notification
                self.refresh_script_list(suppress_notification=True)
                
                # Show notification
                toast = Notification(
                    app_id="PowerShell Script Manager",
                    title="Folder Removed",
                    msg=f"Removed: {folder_path}\n{script_count} script(s) will no longer be included",
                    duration="short"
                )
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            
    def refresh_folder_list(self):
        self.folder_listbox.delete(0, tk.END)
        
        # Store folder paths for reference (needed for context menu actions)
        self.folder_paths = []
        
        for folder in self.app_data.folders:
            # Count scripts in this folder
            script_count = 0
            if os.path.exists(folder):
                for root, _, files in os.walk(folder):
                    script_count += sum(1 for file in files if file.endswith('.ps1'))
            
            # Display folder with script count
            display_text = f"{folder} ({script_count} scripts)"
            self.folder_listbox.insert(tk.END, display_text)
            
            # Store original path for later use
            self.folder_paths.append(folder)
            
    def on_tree_click(self, event):
        tree = event.widget
        item = tree.identify('item', event.x, event.y)
        if not item:
            return
            
        column = tree.identify_column(event.x)
        values = tree.item(item)['values']
        
        if not values or len(values) < 2:
            return
            
        # Clear selection in the other tree view
        other_tree = self.scripts_tree if tree == self.favorites_tree else self.favorites_tree
        other_tree.selection_remove(*other_tree.selection())
            
        script_name = values[1]
        scripts = self.app_data.get_all_powershell_scripts()
        
        for script in scripts:
            if script['name'] == script_name:
                if column == '#1':  # Favorite column
                    self.app_data.toggle_favorite(script['full_path'])
                    self.refresh_script_list(suppress_notification=True)
                else:  # Script name column - show preview
                    try:
                        with open(script['full_path'], 'r', encoding='utf-8') as f:
                            content = f.read()
                            self.preview_text.configure(state='normal')
                            self.preview_text.delete(1.0, tk.END)
                            self.preview_text.insert(tk.END, content)
                            self.preview_text.configure(state='disabled')
                            # Update preview label with script name
                            self.preview_label_frame.configure(text=f"Preview: {script['name']}")
                            # Show action buttons
                            self.show_action_buttons(True)
                    except Exception as e:
                        messagebox.showerror("Error", f"Could not read script: {e}")
                        self.preview_label_frame.configure(text="Script Preview")
                        # Hide action buttons on error
                        self.show_action_buttons(False)
                break

    def refresh_script_list(self, show_startup_notification=False, suppress_notification=False):
        # Store current script count
        current_scripts = set()
        for tree in [self.scripts_tree]:
            for item in tree.get_children():
                values = tree.item(item)['values']
                if values:
                    current_scripts.add(values[1])  # Store script names
        
        # Clear existing items in both trees
        for tree in [self.favorites_tree, self.scripts_tree]:
            for item in tree.get_children():
                tree.delete(item)
            
        # Get and display all scripts
        scripts = self.app_data.get_all_powershell_scripts()
        
        # Sort scripts by name
        scripts.sort(key=lambda x: x['name'].lower())
        
        # Keep track of scripts
        new_scripts = set()
        current_count = len(scripts)
        last_count = self.app_data.last_script_count
        
        for script in scripts:
            heart = '♥' if script['is_favorite'] else '♡'
            values = (heart, script['name'])
            
            # Check if this is a new script
            if script['name'] not in current_scripts:
                new_scripts.add(script['name'])
            
            # Add to appropriate tree(s)
            if script['is_favorite']:
                self.favorites_tree.insert('', 'end', values=values)
            
            # Always add to main script tree
            self.scripts_tree.insert('', 'end', values=values)
        
        # Update the script count
        self.app_data.update_script_count(current_count)
        
        # Show appropriate notification unless suppressed
        if not suppress_notification:
            if show_startup_notification and last_count > 0:
                if current_count > last_count:
                    diff = current_count - last_count
                    msg = f"Found {diff} new script(s) since last session"
                elif current_count < last_count:
                    diff = last_count - current_count
                    msg = f"Missing {diff} script(s) since last session"
                else:
                    msg = "No changes in scripts since last session"
                
                toast = Notification(
                    app_id="PowerShell Script Manager",
                    title="PowerShell Script Count",
                    msg=msg,
                    duration="short"
                )
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            elif not show_startup_notification:
                toast = Notification(
                    app_id="PowerShell Script Manager",
                    title="PowerShell Scan Completed",
                    msg=f"Found {len(new_scripts)} new script(s)" if new_scripts else "No new scripts found",
                    duration="short"
                )
                toast.set_audio(audio.Default, loop=False)
                toast.show()
            
    def setup_system_tray(self):
        try:
            from PIL import Image, ImageTk
            import pystray
            from pystray import MenuItem as item

            # Create menu items
            def create_menu(icon, item):
                if item == 'show':
                    self.show_window()
                elif item == 'exit':
                    self.exit_app()

            # Create system tray icon menu
            menu = (
                item('Show', lambda: create_menu(None, 'show')),
                item('Exit', lambda: create_menu(None, 'exit'))
            )

            # Create a simple icon (white background with black border)
            image = Image.new('RGBA', (64, 64), 'white')
            for x in range(64):
                for y in range(64):
                    if x in [0, 63] or y in [0, 63]:
                        image.putpixel((x, y), (0, 0, 0, 255))

            # Create system tray icon
            self.tray = pystray.Icon("PowerShell Script Manager", image, "PowerShell Script Manager", menu)
            self.tray.run_detached()
        except ImportError:
            messagebox.showwarning("Warning", "System tray feature requires 'pillow' and 'pystray' packages. Install them using: pip install pillow pystray")
    
    def show_window(self):
        self.root.after(0, self.root.deiconify)
        self.root.state('normal')
    
    def hide_window(self):
        self.root.withdraw()
    
    def exit_app(self):
        if hasattr(self, 'tray') and self.tray:
            self.tray.stop()
        self.root.after(0, self.root.destroy)
    
    def start_update_check(self):
        """Start a background thread to check for PowerShell updates"""
        import threading
        
        def check_updates_thread():
            # Check Windows PowerShell update status
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', '(Get-Host).Version.ToString()'],
                    capture_output=True,
                    text=True
                )
                win_ps_version = result.stdout.strip()
                
                # Check if Windows PowerShell update is available (PowerShell 5.1 is the latest for Windows PowerShell)
                current_version = [int(x) for x in win_ps_version.split('.')]
                if len(current_version) >= 2 and (current_version[0] < 5 or (current_version[0] == 5 and current_version[1] < 1)):
                    self.powershell_status['Windows PowerShell'] = {'status': 'update_available', 'version': '5.1'}
                else:
                    self.powershell_status['Windows PowerShell'] = {'status': 'up_to_date', 'version': win_ps_version}
            except:
                self.powershell_status['Windows PowerShell'] = {'status': 'unknown', 'version': 'Unknown'}
            
            # Check PowerShell Core update status
            try:
                result = subprocess.run(
                    ['pwsh', '-Command', '(Get-Host).Version.ToString()'],
                    capture_output=True,
                    text=True
                )
                core_ps_version = result.stdout.strip()
                
                # Check if PowerShell Core update is available via GitHub API
                try:
                    result = subprocess.run(
                        ['powershell.exe', '-Command', 
                         "try { $releaseInfo = Invoke-RestMethod -Uri 'https://api.github.com/repos/PowerShell/PowerShell/releases/latest' -TimeoutSec 3; $releaseInfo.tag_name.TrimStart('v') } catch { 'Unknown' }"],
                        capture_output=True,
                        text=True,
                        timeout=5
                    )
                    latest_version = result.stdout.strip()
                    
                    if latest_version != 'Unknown':
                        current_parts = [int(x) for x in core_ps_version.split('.')]
                        latest_parts = [int(x) for x in latest_version.split('.')]
                        
                        # Compare versions
                        update_available = False
                        for i in range(min(len(current_parts), len(latest_parts))):
                            if latest_parts[i] > current_parts[i]:
                                update_available = True
                                break
                            elif current_parts[i] > latest_parts[i]:
                                break
                        
                        if update_available:
                            self.powershell_status['PowerShell Core'] = {'status': 'update_available', 'version': latest_version}
                        else:
                            self.powershell_status['PowerShell Core'] = {'status': 'up_to_date', 'version': core_ps_version}
                    else:
                        self.powershell_status['PowerShell Core'] = {'status': 'unknown', 'version': core_ps_version}
                except:
                    self.powershell_status['PowerShell Core'] = {'status': 'unknown', 'version': core_ps_version}
            except:
                self.powershell_status['PowerShell Core'] = {'status': 'not_installed', 'version': 'N/A'}
            
            # Check PowerShell ISE
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', "if (Get-Command powershell_ise.exe -ErrorAction SilentlyContinue) { (Get-Host).Version.ToString() } else { 'Not installed' }"],
                    capture_output=True,
                    text=True
                )
                ise_version = result.stdout.strip()
                
                if ise_version != 'Not installed':
                    # ISE version follows Windows PowerShell version
                    if self.powershell_status.get('Windows PowerShell', {}).get('status') == 'update_available':
                        self.powershell_status['PowerShell ISE'] = {'status': 'update_available', 'version': '5.1'}
                    else:
                        self.powershell_status['PowerShell ISE'] = {'status': 'up_to_date', 'version': ise_version}
                else:
                    self.powershell_status['PowerShell ISE'] = {'status': 'not_installed', 'version': 'N/A'}
            except:
                self.powershell_status['PowerShell ISE'] = {'status': 'unknown', 'version': 'Unknown'}
            
            # Check PowerShell Preview
            try:
                result = subprocess.run(
                    ['powershell.exe', '-Command', 
                     "if (Get-Command pwsh-preview -ErrorAction SilentlyContinue) { pwsh-preview -Command '(Get-Host).Version.ToString()' } else { 'Not installed' }"],
                    capture_output=True,
                    text=True
                )
                preview_version = result.stdout.strip()
                
                if preview_version != 'Not installed':
                    # Preview versions are typically already the latest
                    self.powershell_status['PowerShell Preview'] = {'status': 'preview', 'version': preview_version}
                else:
                    self.powershell_status['PowerShell Preview'] = {'status': 'not_installed', 'version': 'N/A'}
            except:
                self.powershell_status['PowerShell Preview'] = {'status': 'unknown', 'version': 'Unknown'}
            
            # Signal that the update check is complete
            self.root.event_generate('<<UpdateCheckComplete>>', when='tail')
        
        # Start the update check in a background thread
        update_thread = threading.Thread(target=check_updates_thread)
        update_thread.daemon = True
        update_thread.start()
        
        # Bind event to update UI when check is complete
        self.root.bind('<<UpdateCheckComplete>>', lambda e: self.update_powershell_ui())
    
    def update_powershell_ui(self):
        """Update the PowerShell tab UI with the latest status information"""
        # Only update if the PowerShell tab has been initialized
        if hasattr(self, 'powershell_tab') and self.powershell_tab.winfo_exists():
            self.refresh_powershell_tab()
    
    def refresh_powershell_tab(self):
        """Refresh the PowerShell tab with current information"""
        # Clear the tab
        for widget in self.powershell_tab.winfo_children():
            widget.destroy()
        
        # Re-setup the tab
        self.setup_powershell_tab()
        
    def load_icons(self):
        self.icons = {}
        
        if not PILLOW_AVAILABLE:
            return
            
        # Base64 encoded small icons (16x16)
        icons_data = {
            'refresh': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAA7UlEQVQ4je
                3SMUoDQRTG8d+uG0ijjQhewANY2VtY2Ai2gljoEQRbzyBYegC9gHdIYyWCYGUjNhYWNiLimsxYJLPZjVkV
                /Jrh/d/78GYYIrIW5G7xgDs8Yfk/8B6mj/Em+7Y59AYfsYfzKniBKU4wLeKHOPwNfluB21jHKhZxhtOf5h
                Y6EZ3qM/cwxZeY48Y3eJq42sU2RrjCEB2sYIw5LjFFP/4j9EqSJXooQF84xxw32MQrNnCEl9DvR7yINdoh
                9nGJuziw43iPjQ3c4QHHOMMoiJr4qOwEXbxjOyLGMXbntQpqA5rQrALXbfMDImFEyWiHZSkAAAAASUVORK
                5CYII=
            ''',
            'notepad': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAAdklEQVQ4jWNg
                oDH4TwGmngEsJKj5j6aGYgMYGBgYmHAI/kdisZHmAoAAZsB/BvJduPAGsJBjAAsSzYRFw7JjZxm+/fiJVYO+
                lirDkgUz4S5gIeQFXJqxsUEuYMKnCB3gs4iFmIDE5gIQICoaR8PAGIAvHMjSAwDT+xX4xB8kWAAAAABJRU5E
                rkJggg==
            ''',
            'run': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAArklEQVQ4jc2S
                sQ3CMBBF31EKGjpnADagYgQmoGICxAZkBDagYwMydAgKKrCLRCIjRQLxJUu+//25s0+60b5iB5RAdq80pyXw
                PjGLc58ZKT6+JDYDLTAAu583lRngDLxcPgAbH6S1GwIPl9TAHdirpgPORm+AnSO0wAV4m/kKeKkMG1Qfk4BD
                gP4IPIClOcPacoTRVkXodWLj/JkkQV/rOV2mTwjThlCrnXcSdNSUXhIfjeTNejQc7PQAAAAASUVORK5CYII=
            ''',
            'run_as': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAA60lEQVQ4je2R
                MUvDUBSFv/uSFkQoFReLQzfBSXB0UxH/gbj6H/wPDg4OTg5OFRdx1cVBnJUOhYrg4iQUaYPa9zgkfdVXOokH
                Lu/de8+5vHMh5YTkLtkbHgBPQOR5QJ6ZMcDJ2K/fQYEEWAO+gRPgRaTZr7nQS/pWgHPgGzgFLoB74BLYBe6A
                SgJ4FdEcmAXKwCkwQPOvAdshGA3KQLTnRPQAqOOmIdR27gGTIhAnuWdUwU2fARogE4kaLnPWrSqiidf2JtB1
                UhfYEikk1bQH7HuNK86uAN+BWewDK1OmwDqwC/wASzkHnTnxH1qyF46DPsnXAAAAAElFTkSuQmCC
            ''',
            'folder_add': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAApElEQVQ4je3S
                MQrCQBQE0JdoYSNYCDmA5ATmJt7AtBZeKL2FjQiijWghiI2F2Aie4A2MYGIRFhIWQiCF+Zt/mJ3Z2R/8C1FM
                cUIHi6AF+7hihi7GIQRdDPAMUrCCKzboRQ4iXSzxwDlWkGHtRcFGs/AsalwS88iNSnXdx7is8x09XLBPFWzR
                xx33VMEBGfa4pQoOaGOESYogwxxTLGInbUWxwzFxMokszqhqKZHpAAAAAElFTkSuQmCC
            ''',
            'folder_remove': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAAX0lEQVQ4je3S
                wQnAIAxG4T/qSAfqSB1JR+pI3aSblOIFD4KHFkI9vFNICF8S4PM4YTOwiojnN4FiY8CKHXdEPOkEGXO9oKVz
                JcjxhL0myCt/MeBoJRjr8gVR2v9J8B+CG9voHdk+Jh6BAAAAAElFTkSuQmCC
            ''',
            'folder_list': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAAlklEQVQ4jWNg
                gIL///8z/v//n5FYNUzEGMpIrGZcBvz//58FXQMjIyMTOpsRmY3NcJxeQNfMwMDAwIRNM7JmZD4MsOBzPjLY
                snM3w7cfP7FqiA0KZEhPSUa4AKcX8GkGA5wG4NOMDXAZwEKs6cguQPYKTgOwacamGQawMDAwMLz9+JGBkZGR
                gYWZmeEfijwjw9uPHwkaCACyXyPGY23BSAAAAABJRU5ErkJggg==
            ''',
            'scripts_refresh': '''
                iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAABmJLR0QA/wD/AP+gvaeTAAAAo0lEQVQ4je3R
                MQrCUBCF4S8qFjYWQdBTWNhYCILgATyBpZ1H8BSWXsBGvIkXsLLRSrCyUAgphMeWhUTxbZoUFgMzu/vP7OzC
                vy5wwh5dVIodYYNnSl3kPUPge8l7ung1EAwwyqBDrDPowB8FexyxwTqD7pZO0MYyt6CKMWZ+XXotNSNnYF7A
                6AXKJQQr9DN6xSJBeZhOnDVmfqVpr6UCG6ySwBNlWTvHgW2nCgAAAABJRU5ErkJggg==
            '''
        }
        
        # Create PhotoImage objects from base64 data
        for name, data in icons_data.items():
            try:
                # Clean up the base64 data
                clean_data = ''.join(line.strip() for line in data.splitlines())
                # Decode the base64 data
                image_data = base64.b64decode(clean_data)
                # Create an image from the data
                image = Image.open(BytesIO(image_data))
                # Convert to PhotoImage for tkinter
                self.icons[name] = ImageTk.PhotoImage(image)
            except Exception as e:
                print(f"Failed to load icon {name}: {e}")
    
    def run_script_as(self):
        # Get the currently selected item from either tree
        selected_script_path = self.get_selected_script_path()
        
        if not selected_script_path:
            messagebox.showinfo("Run As...", "No script selected.")
            return
        
        # Check execution policy before attempting to run
        # Note: We still allow "Run As..." even with restricted policy as it might be used to change the policy
        if not self.check_execution_policy():
            response = messagebox.askyesno("Execution Policy Warning", 
                "PowerShell execution policy is set to restrict script execution. \n\n"
                "Running as administrator might still work if you change the execution policy.\n\n"
                "Do you want to continue?")
            if not response:
                return
        
        try:
            if sys.platform != 'win32':
                messagebox.showerror("Error", "Run As Administrator is only supported on Windows.")
                return
                
            # Use ShellExecute to run PowerShell with admin privileges
            if hasattr(ctypes, 'windll'):
                # The command to run the PowerShell script
                powershell_exe = "powershell.exe"
                
                # Use ShellExecute to run PowerShell as administrator with the script
                ctypes.windll.shell32.ShellExecuteW(
                    None, 
                    "runas",  # Run as administrator
                    powershell_exe,
                    f"-File \"{selected_script_path}\"",  # Pass the script as a parameter
                    os.path.dirname(selected_script_path),  # Set working directory to script location
                    1  # SW_SHOWNORMAL
                )
            else:
                messagebox.showerror("Error", "Could not access Windows API for elevated privileges.")
        except Exception as e:
            messagebox.showerror("Error", f"Could not run script with admin privileges: {e}")

    def run_script(self):
        # Get the currently selected item from either tree
        selected_script_path = self.get_selected_script_path()
        
        if not selected_script_path:
            messagebox.showinfo("Run", "No script selected.")
            return
            
        # Check execution policy before attempting to run
        if not self.check_execution_policy():
            messagebox.showerror("Execution Policy Error", 
                "PowerShell execution policy is set to restrict script execution. \n\n"
                "To change this, open PowerShell as Administrator and run:\n"
                "Set-ExecutionPolicy -ExecutionPolicy RemoteSigned\n\n"
                "For more information, see about_Execution_Policies at:\n"
                "https://go.microsoft.com/fwlink/?LinkID=135170")
            return
        
        try:
            # Run the script using PowerShell
            subprocess.Popen(['powershell.exe', '-File', selected_script_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not run script: {e}")
            
    def get_selected_script_path(self):
        # Get the currently selected item from either tree
        selected_item = None
        selected_tree = None
        
        for tree in [self.favorites_tree, self.scripts_tree]:
            selection = tree.selection()
            if selection:
                selected_item = selection[0]
                selected_tree = tree
                break
        
        if not selected_item or not selected_tree:
            return None
        
        # Get the script name
        values = selected_tree.item(selected_item)['values']
        if not values or len(values) < 2:
            return None
            
        script_name = values[1]
        
        # Find the script in the app data
        scripts = self.app_data.get_all_powershell_scripts()
        for script in scripts:
            if script['name'] == script_name:
                return script['full_path']
                
        return None

    def open_in_notepad(self):
        # Get the currently selected item from either tree
        selected_script_path = self.get_selected_script_path()
        
        if not selected_script_path:
            messagebox.showinfo("Open in Notepad", "No script selected.")
            return
        
        try:
            # Open the script in Notepad
            subprocess.Popen(['notepad.exe', selected_script_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open script in Notepad: {e}")

    def refresh_preview(self):
        # Get the currently selected item from either tree
        selected_script_path = self.get_selected_script_path()
        
        if not selected_script_path:
            messagebox.showinfo("Refresh Preview", "No script selected.")
            return
            
        # Find script name to update preview label
        script_name = os.path.basename(selected_script_path)
        
        try:
            # Reload the script content
            with open(selected_script_path, 'r', encoding='utf-8') as f:
                content = f.read()
                self.preview_text.configure(state='normal')
                self.preview_text.delete(1.0, tk.END)
                self.preview_text.insert(tk.END, content)
                self.preview_text.configure(state='disabled')
                self.preview_label_frame.configure(text=f"Preview: {script_name}")
                # Ensure action buttons are visible when refreshing
                self.show_action_buttons(True)
        except Exception as e:
            messagebox.showerror("Error", f"Could not refresh script preview: {e}")
            # Hide action buttons on error
            self.show_action_buttons(False)
    
    def show_folder_context_menu(self, event):
        # First select the item under the cursor
        try:
            index = self.folder_listbox.nearest(event.y)
            if index >= 0:
                self.folder_listbox.selection_clear(0, tk.END)
                self.folder_listbox.selection_set(index)
                self.folder_listbox.activate(index)
                
                # Display the context menu
                try:
                    self.folder_context_menu.tk_popup(event.x_root, event.y_root)
                finally:
                    self.folder_context_menu.grab_release()
        except Exception as e:
            print(f"Error showing context menu: {e}")
    
    def open_folder_location(self):
        selection = self.folder_listbox.curselection()
        if selection:
            # Get the actual folder path from our stored list
            index = selection[0]
            if index < len(self.folder_paths):
                folder_path = self.folder_paths[index]
                try:
                    if os.path.exists(folder_path):
                        # Open folder in Windows Explorer
                        os.startfile(folder_path)
                    else:
                        messagebox.showerror("Error", f"Folder does not exist: {folder_path}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not open folder: {e}")
    
    def on_script_right_click(self, event):
        # Identify the tree that was clicked
        tree = event.widget
        
        # Get the item under the cursor
        item = tree.identify_row(event.y)
        if not item:
            return
        
        # Select the item
        tree.selection_set(item)
        
        # Clear selection in the other tree
        other_tree = self.scripts_tree if tree == self.favorites_tree else self.favorites_tree
        other_tree.selection_remove(*other_tree.selection())
        
        # Show context menu
        try:
            self.script_context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.script_context_menu.grab_release()
    
    def show_action_buttons(self, show=True):
        """Show or hide the action buttons based on whether a script is selected"""
        # Remove all buttons first
        for widget in self.action_buttons_frame.winfo_children():
            widget.pack_forget()
        
        # If show is True, show the buttons
        if show:
            self.action_buttons_frame.pack(side='right')
            self.refresh_btn.pack(side='left', padx=(0,5))
            self.notepad_btn.pack(side='left', padx=(0,5))
            self.run_btn.pack(side='left', padx=(0,5))
            self.runas_btn.pack(side='left', padx=(0,5))
        else:
            self.action_buttons_frame.pack_forget()
            
    def toggle_script_favorite(self):
        # Get the currently selected item from either tree
        selected_item = None
        selected_tree = None
        
        for tree in [self.favorites_tree, self.scripts_tree]:
            selection = tree.selection()
            if selection:
                selected_item = selection[0]
                selected_tree = tree
                break
        
        if not selected_item or not selected_tree:
            return
        
        # Get the script name
        values = selected_tree.item(selected_item)['values']
        if not values or len(values) < 2:
            return
            
        script_name = values[1]
        
        # Find the script in the app data
        scripts = self.app_data.get_all_powershell_scripts()
        for script in scripts:
            if script['name'] == script_name:
                # Toggle the favorite status
                self.app_data.toggle_favorite(script['full_path'])
                self.refresh_script_list(suppress_notification=True)
                break
        
    def on_tree_select(self, event):
        tree = event.widget
        selection = tree.selection()
        
        if not selection:
            return
            
        # Clear selection in the other tree
        other_tree = self.scripts_tree if tree == self.favorites_tree else self.favorites_tree
        other_tree.selection_remove(*other_tree.selection())
        
        # Get the selected item
        item = selection[0]
        values = tree.item(item)['values']
        
        if not values or len(values) < 2:
            return
            
        script_name = values[1]
        scripts = self.app_data.get_all_powershell_scripts()
        
        # Show preview of selected script
        for script in scripts:
            if script['name'] == script_name:
                try:
                    with open(script['full_path'], 'r', encoding='utf-8') as f:
                        content = f.read()
                        self.preview_text.configure(state='normal')
                        self.preview_text.delete(1.0, tk.END)
                        self.preview_text.insert(tk.END, content)
                        self.preview_text.configure(state='disabled')
                        # Update the LabelFrame text directly
                        self.preview_label_frame.configure(text=f"Preview: {script['name']}")
                        # Show action buttons
                        self.show_action_buttons(True)
                except Exception as e:
                    messagebox.showerror("Error", f"Could not read script: {e}")
                    self.preview_label_frame.configure(text="Script Preview")
                    # Hide action buttons on error
                    self.show_action_buttons(False)
                break

def main():
    root = tk.Tk()
    app = PowerShellManager(root)
    root.mainloop()

if __name__ == "__main__":
    main()
