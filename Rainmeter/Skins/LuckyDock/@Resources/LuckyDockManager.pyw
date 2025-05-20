import tkinter as tk
from tkinter import messagebox, font, filedialog, ttk
import configparser
import os
import subprocess
import re
import sys

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

if sys.platform.startswith('win'):
    try:
        import win32gui
        import win32con

        if not sys.executable.endswith('pythonw.exe') and not hasattr(sys, 'frozen'):
            hwnd = win32gui.GetForegroundWindow()
            win32gui.ShowWindow(hwnd, win32con.SW_HIDE)
    except ImportError:
        import logging
        logging.warning("win32gui not available — consider renaming the file to .pyw to avoid console.")


class LuckyDockManager:
    def __init__(self, master, dock_number=None):
        self.master = master
        self.master.title("LuckyDockManager")
        self.master.geometry("1800x1300")

        self.bg_color = "#1e1e1e"
        self.fg_color = "#ffffff"
        self.entry_bg = "#2d2d2d"
        self.button_bg = "#3e3e3e"
        self.button_hover_bg = "#4a4a4a"
        self.listbox_bg = "#252525"
        self.selected_bg = "#0078d7"
        self.status_bg = "#1a1a1a"
        self.accent_color = "#007acc"
        self.border_color = "#3e3e3e"
        self.separator_color = "#888888"
        
        self.selected_dock_border = "#ffffff"
        
        self.master.configure(bg=self.bg_color)
        self.master.option_add('*TButton*highlightBackground', self.bg_color)
        self.master.option_add('*TButton*highlightColor', self.bg_color)
        
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.base_dir = os.path.dirname(script_dir)
        
        self.available_docks = self.find_available_docks()
        
        self.current_dock = tk.StringVar()
        if self.available_docks:
            if dock_number is not None:
                dock_name = f"LuckyDock{dock_number}"
                if dock_name in self.available_docks:
                    self.current_dock.set(dock_name)
                else:
                    self.current_dock.set(self.available_docks[0])
            else:
                self.current_dock.set(self.available_docks[0])
        
        self.ini_file = None
        self.entries = []
        self.selected_entry_index = None
        self.status_message = tk.StringVar()
        self.status_message.set("Ready")
        
        self.create_widgets()
        
        self.setup_keyboard_shortcuts()
        
        if self.available_docks:
            self.load_dock(self.current_dock.get())

    def find_available_docks(self):
        docks = []
        
        for item in os.listdir(self.base_dir):
            if item.startswith("Dock") and os.path.isdir(os.path.join(self.base_dir, item)):
                dock_dir = os.path.join(self.base_dir, item)
                for file in os.listdir(dock_dir):
                    if file.startswith("LuckyDock") and file.endswith(".ini"):
                        dock_name = os.path.splitext(file)[0]
                        docks.append(dock_name)
        
        return sorted(docks)

    def create_widgets(self):
        self.style = ttk.Style()
        self.style.theme_use('default')
        
        self.style.configure('TFrame', background=self.bg_color)
        self.style.configure('TLabel', background=self.bg_color, foreground=self.fg_color, font=('Segoe UI', 10))
        
        self.style.configure('TButton', 
                            background=self.button_bg, 
                            foreground=self.fg_color,
                            font=('Segoe UI', 10),
                            borderwidth=0,
                            padding=6,
                            relief=tk.FLAT)
        self.style.map('TButton', 
                      background=[('active', self.button_hover_bg)],
                      relief=[('pressed', tk.SUNKEN)])
        
        self.style.configure('Accent.TButton', 
                            background=self.accent_color, 
                            foreground=self.fg_color,
                            font=('Segoe UI', 10, 'bold'),
                            padding=6)
        self.style.map('Accent.TButton', 
                      background=[('active', '#0069b3')])
        
        self.style.configure('TEntry', 
                            fieldbackground=self.entry_bg, 
                            foreground=self.fg_color,
                            borderwidth=1,
                            relief=tk.SOLID)
        self.style.configure('TCombobox', 
                            fieldbackground=self.entry_bg, 
                            foreground=self.fg_color,
                            background=self.button_bg,
                            arrowcolor=self.fg_color)
        self.style.map('TCombobox', 
                      fieldbackground=[('readonly', self.entry_bg)],
                      selectbackground=[('readonly', self.selected_bg)])
        
        self.style.configure('Title.TLabel', background=self.bg_color, foreground=self.fg_color, font=('Segoe UI', 16, 'bold'))
        self.style.configure('FieldLabel.TLabel', background=self.bg_color, foreground=self.fg_color, font=('Segoe UI', 12, 'bold'))
        
        self.style.configure('DockSelector.TCombobox',
                         font=('Segoe UI', 16, 'bold'),
                         padding=8)
        self.style.map('DockSelector.TCombobox',
                    fieldbackground=[('readonly', self.entry_bg)],
                    selectbackground=[('readonly', self.selected_bg)],
                    selectforeground=[('readonly', self.fg_color)])
        
        self.style.configure('DockBox.TFrame', 
                         background=self.bg_color,
                         borderwidth=1,
                         relief=tk.SOLID,
                         bordercolor=self.selected_dock_border)
        
        self.style.configure('DockLabel.TLabel',
                         background=self.bg_color,
                         foreground=self.fg_color,
                         font=('Segoe UI', 14, 'bold'))
        
        main_container = ttk.Frame(self.master, style='TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        title_frame = ttk.Frame(main_container, style='TFrame')
        title_frame.pack(pady=10, fill=tk.X)
        
        try:
            title_font = font.Font(family="Impact", size=36, weight="bold")
        except tk.TclError:
            try:
                title_font = font.Font(family="Verdana", size=34, weight="bold")
            except tk.TclError:
                title_font = font.Font(size=36, weight="bold")
        
        subtitle_font = self.get_font(14)
        
        title_container = ttk.Frame(title_frame, style='TFrame')
        title_container.pack(pady=5)
        
        title_label = ttk.Label(title_container, text="LuckyDockManager", 
                 font=title_font, foreground="#00aaff", style='TLabel')
        title_label.pack()
        
        ttk.Separator(main_container, orient='horizontal').pack(fill=tk.X, pady=10)
        
        dock_frame = ttk.Frame(main_container, style='TFrame')
        dock_frame.pack(pady=20, fill=tk.X)
        
        center_container = ttk.Frame(dock_frame, style='TFrame')
        center_container.pack(fill=tk.X)
        
        left_spacer = ttk.Frame(center_container, style='TFrame')
        left_spacer.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        dock_selector_container = ttk.Frame(center_container, style='DockBox.TFrame', padding=(15, 12))
        dock_selector_container.pack(side=tk.LEFT, pady=(5, 15))
        
        right_spacer = ttk.Frame(center_container, style='TFrame')
        right_spacer.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        editing_label = ttk.Label(dock_selector_container, 
                                text="Selected Dock:", 
                                font=self.get_font(15, "bold"), 
                                style='DockLabel.TLabel')
        editing_label.pack(pady=(0, 10), anchor=tk.CENTER)
        
        dock_combo = ttk.Combobox(dock_selector_container, 
                                textvariable=self.current_dock, 
                                values=self.available_docks, 
                                state="readonly", 
                                style='DockSelector.TCombobox', 
                                width=20,
                                height=25)
        dock_combo.pack()
        dock_combo.bind("<<ComboboxSelected>>", self.on_dock_selected)
        
        content_frame = ttk.Frame(main_container, style='TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, pady=15)
        
        left_frame = ttk.Frame(content_frame, style='TFrame', padding=(0, 0, 15, 0))
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        ttk.Label(left_frame, text="Dock Entries", 
                 font=self.get_font(16, "bold"), style='Title.TLabel').pack(anchor=tk.W, pady=(0, 15))
        
        entries_frame = ttk.Frame(left_frame, style='TFrame')
        entries_frame.pack(fill=tk.BOTH, expand=True)
        
        self.entries_listbox = tk.Listbox(entries_frame, 
                                         bg=self.listbox_bg, 
                                         fg=self.fg_color, 
                                         selectbackground=self.selected_bg, 
                                         font=self.get_font(16),
                                         borderwidth=1, 
                                         relief=tk.SOLID,
                                         activestyle='none',
                                         highlightthickness=0)
        self.entries_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.entries_listbox.bind('<<ListboxSelect>>', self.on_select_entry)
        
        scrollbar = ttk.Scrollbar(entries_frame, orient=tk.VERTICAL, command=self.entries_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.entries_listbox.config(yscrollcommand=scrollbar.set)
        
        list_buttons_frame = ttk.Frame(left_frame, style='TFrame')
        list_buttons_frame.pack(fill=tk.X, pady=10)
        
        top_buttons = ttk.Frame(list_buttons_frame, style='TFrame')
        top_buttons.pack(fill=tk.X, pady=(0, 5))
        
        add_button = ttk.Button(top_buttons, text="Add Entry", command=self.add_entry)
        add_button.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        add_separator_button = ttk.Button(top_buttons, text="Add Separator", command=self.add_separator)
        add_separator_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        middle_buttons = ttk.Frame(list_buttons_frame, style='TFrame')
        middle_buttons.pack(fill=tk.X, pady=(0, 5))
        
        remove_button = ttk.Button(middle_buttons, text="Remove Entry", command=self.remove_entry)
        remove_button.pack(fill=tk.X, expand=True)
        
        bottom_buttons = ttk.Frame(list_buttons_frame, style='TFrame')
        bottom_buttons.pack(fill=tk.X)
        
        move_up_button = ttk.Button(bottom_buttons, text="Move Up", command=lambda: self.move_entry(-1))
        move_up_button.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)
        
        move_down_button = ttk.Button(bottom_buttons, text="Move Down", command=lambda: self.move_entry(1))
        move_down_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Separator(content_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=15)
        
        right_frame = ttk.Frame(content_frame, style='TFrame')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        ttk.Label(right_frame, text="Entry Details", 
                 font=self.get_font(16, "bold"), style='Title.TLabel').pack(anchor=tk.W, pady=(0, 15))
        
        details_frame = ttk.Frame(right_frame, style='TFrame')
        details_frame.pack(fill=tk.BOTH, expand=True)
        
        name_frame = ttk.Frame(details_frame, style='TFrame')
        name_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(name_frame, text="Name:", width=12, style='FieldLabel.TLabel').pack(side=tk.LEFT)
        self.name_entry = ttk.Entry(name_frame, font=self.get_font(12))
        self.name_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        
        app_path_frame = ttk.Frame(details_frame, style='TFrame')
        app_path_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(app_path_frame, text="App Path:", width=12, style='FieldLabel.TLabel').pack(side=tk.LEFT)
        self.app_path_entry = ttk.Entry(app_path_frame, font=self.get_font(12))
        self.app_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        self.browse_app_button = ttk.Button(app_path_frame, text="Browse", command=self.browse_app_path, width=10)
        self.browse_app_button.pack(side=tk.RIGHT)
        
        icon_path_frame = ttk.Frame(details_frame, style='TFrame')
        icon_path_frame.pack(fill=tk.X, pady=10)
        
        ttk.Label(icon_path_frame, text="Icon Path:", width=12, style='FieldLabel.TLabel').pack(side=tk.LEFT)
        self.icon_path_entry = ttk.Entry(icon_path_frame, font=self.get_font(12))
        self.icon_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        self.browse_icon_button = ttk.Button(icon_path_frame, text="Browse", command=self.browse_icon_path, width=10)
        self.browse_icon_button.pack(side=tk.RIGHT)
        
        update_frame = ttk.Frame(details_frame, style='TFrame')
        update_frame.pack(fill=tk.X, pady=(20, 10))
        
        self.update_button = ttk.Button(update_frame, text="Update Entry", command=self.update_entry, 
                  style='Accent.TButton', padding=10)
        self.update_button.pack(fill=tk.X)
        
        ttk.Separator(main_container, orient='horizontal').pack(fill=tk.X, pady=15)
        
        bottom_frame = ttk.Frame(main_container, style='TFrame')
        bottom_frame.pack(fill=tk.X)
        
        save_button = ttk.Button(bottom_frame, text="Save Changes", command=self.save_ini, 
                               style='Accent.TButton', padding=10)
        save_button.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        refresh_button = ttk.Button(bottom_frame, text="Refresh Dock", command=self.refresh_rainmeter, 
                                  padding=10)
        refresh_button.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        status_frame = ttk.Frame(self.master, style='TFrame')
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        status_bar = ttk.Label(status_frame, textvariable=self.status_message, 
                              background=self.status_bg, foreground="#aaaaaa",
                              relief=tk.FLAT, anchor=tk.W, padding=(10, 5))
        status_bar.pack(fill=tk.X)

    def get_font(self, size, weight="normal"):
        try:
            return font.Font(family="Segoe UI", size=size, weight=weight)
        except tk.TclError:
            return font.Font(size=size, weight=weight)

    def setup_keyboard_shortcuts(self):
        self.master.bind("<Control-s>", lambda event: self.save_ini())
        self.master.bind("<Control-n>", lambda event: self.add_entry())
        self.master.bind("<Control-d>", lambda event: self.remove_entry())
        self.master.bind("<Control-Up>", lambda event: self.move_entry(-1))
        self.master.bind("<Control-Down>", lambda event: self.move_entry(1))

    def set_status(self, message):
        self.status_message.set(message)
        self.master.update_idletasks()

    def on_dock_selected(self, event):
        selected_dock = self.current_dock.get()
        self.load_dock(selected_dock)

    def load_dock(self, dock_name):
        self.set_status(f"Loading {dock_name}...")
        
        dock_number = None
        if dock_name.startswith("LuckyDock") and dock_name[9:].isdigit():
            dock_number = dock_name[9:]
        
        if dock_number:
            self.ini_file = os.path.join(self.base_dir, f"Dock{dock_number}", f"{dock_name}.ini")
            
            self.load_ini()
        else:
            messagebox.showerror("Error", f"Invalid dock name format: {dock_name}")
            self.set_status(f"Error: Invalid dock name format: {dock_name}")

    def browse_app_path(self):
        file_path = filedialog.askopenfilename(
            title="Select Application",
            filetypes=[("Executable files", "*.exe"), ("All files", "*.*")]
        )
        if file_path:
            self.app_path_entry.delete(0, tk.END)
            self.app_path_entry.insert(0, f'"{file_path}"')
            self.set_status(f"Selected application: {file_path}")

    def browse_icon_path(self):
        file_path = filedialog.askopenfilename(
            title="Select Icon",
            filetypes=[
                ("Image files", "*.png *.jpg *.jpeg *.ico *.bmp *.gif"),
                ("All files", "*.*")
            ]
        )
        if file_path:
            self.icon_path_entry.delete(0, tk.END)
            self.icon_path_entry.insert(0, file_path)
            self.set_status(f"Selected icon: {file_path}")

    def load_ini(self):
        self.set_status(f"Loading {self.ini_file}...")
        
        if not os.path.exists(self.ini_file):
            messagebox.showerror("Error", f"Could not find {self.ini_file}")
            self.set_status(f"Error: Could not find {self.ini_file}")
            return

        self.entries = []
        self.entries_listbox.delete(0, tk.END)

        try:
            config = configparser.ConfigParser()
            config.read(self.ini_file)

            excluded_sections = ["Variables", "Rainmeter", "Metadata", 
                                "MeasureAppCount", "Style", "MeasureFlash", 
                                "BackgroundMeter"]
            
            for section in config.sections():
                if any(section.startswith(prefix) for prefix in excluded_sections):
                    continue
                
                name = section
                app_path = self.extract_app_path(config[section].get("LeftMouseUpAction", ""))
                icon_path = config[section].get("ImageName", "")
                
                is_separator = name.startswith("Separator_")
                
                entry = {
                    "name": name, 
                    "app_path": app_path, 
                    "icon_path": icon_path,
                    "is_separator": is_separator
                }
                
                self.entries.append(entry)
                self.entries_listbox.insert(tk.END, name)
                
                if is_separator:
                    self.entries_listbox.itemconfig(tk.END, fg=self.separator_color)
            
            self.set_status(f"Loaded {len(self.entries)} entries from {self.ini_file}")
            
            self.selected_entry_index = None
            self.name_entry.delete(0, tk.END)
            self.app_path_entry.delete(0, tk.END)
            self.icon_path_entry.delete(0, tk.END)
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load INI file: {str(e)}")
            self.set_status(f"Error loading INI file: {str(e)}")

    def extract_app_path(self, action_string):
        match = re.search(r'\["([^"]*)"\]', action_string)
        return f'"{match.group(1)}"' if match else ""

    def save_ini(self):
        if not self.ini_file:
            messagebox.showerror("Error", "No dock selected")
            self.set_status("Error: No dock selected")
            return
            
        self.set_status("Saving changes...")
        
        if not self.validate_entries():
            self.set_status("Save cancelled due to validation errors")
            return
        
        try:
            config = configparser.ConfigParser()
            config.read(self.ini_file)

            excluded_sections = ["Variables", "Rainmeter", "Metadata", 
                                "MeasureAppCount", "Style", "MeasureFlash", 
                                "BackgroundMeter"]
            
            sections_to_remove = [
                section for section in config.sections() 
                if not any(section.startswith(prefix) for prefix in excluded_sections)
            ]
            
            for section in sections_to_remove:
                config.remove_section(section)

            for i, entry in enumerate(self.entries):
                section_name = entry["name"]
                config[section_name] = {}
                config[section_name]["Meter"] = "Image"
                config[section_name]["MeterStyle"] = "Style"
                config[section_name]["ImageName"] = entry["icon_path"]
                
                config[section_name]["X"] = "((#BackgroundWidth# - #IconSize#) / 2)"
                config[section_name]["Y"] = f"(#PadTop# + {i} * (#IconSize# + #VerticalGap#))"
                
                if entry.get("is_separator", False):
                    config[section_name]["DynamicVariables"] = "1"
                else:
                    config[section_name]["LeftMouseUpAction"] = f'[{entry["app_path"]}][!SetVariable CurrentIcon "{section_name}"][!CommandMeasure MeasureFlash "Execute 1"]'
                    config[section_name]["MouseOverAction"] = '[!SetOption #CURRENTSECTION# W "(#IconSize# + #HoverSize#)"][!SetOption #CURRENTSECTION# H "(#IconSize# + #HoverSize#)"][!SetOption #CURRENTSECTION# X "([#CURRENTSECTION#:X] - (#HoverSize# / 2))"][!SetOption #CURRENTSECTION# Y "([#CURRENTSECTION#:Y] - (#HoverSize# / 2))"][!UpdateMeter #CURRENTSECTION#][!Redraw]'
                    config[section_name]["MouseLeaveAction"] = f'[!SetOption #CURRENTSECTION# W "#IconSize#"][!SetOption #CURRENTSECTION# H "#IconSize#"][!SetOption #CURRENTSECTION# X "((#BackgroundWidth# - #IconSize#) / 2)"][!SetOption #CURRENTSECTION# Y "(#PadTop# + {i} * (#IconSize# + #VerticalGap#))"][!UpdateMeter #CURRENTSECTION#][!Redraw]'
                    config[section_name]["DynamicVariables"] = "1"

            config["Variables"]["AppCount"] = str(len(self.entries))

            with open(self.ini_file, 'w') as configfile:
                config.write(configfile)

            self.refresh_rainmeter()
            messagebox.showinfo("Success", "INI file saved and Rainmeter skin refreshed")
            self.set_status("Changes saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save INI file: {str(e)}")
            self.set_status(f"Error saving INI file: {str(e)}")

    def validate_entries(self):
        for i, entry in enumerate(self.entries):
            if not entry["name"].strip():
                messagebox.showerror("Validation Error", f"Entry #{i+1} has an empty name")
                return False
            
            name_count = sum(1 for e in self.entries if e["name"] == entry["name"])
            if name_count > 1:
                messagebox.showerror("Validation Error", f"Duplicate entry name: {entry['name']}")
                return False
            
            if not entry.get("is_separator", False):
                if not entry["app_path"].strip() or entry["app_path"] == '""':
                    messagebox.showwarning("Warning", f"Entry '{entry['name']}' has an empty app path")
                    if not messagebox.askyesno("Confirm", "Some entries have empty app paths. Save anyway?"):
                        return False
            
            if not entry["icon_path"].strip() or entry["icon_path"] == '""':
                messagebox.showwarning("Warning", f"Entry '{entry['name']}' has an empty icon path")
                if not messagebox.askyesno("Confirm", "Some entries have empty icon paths. Save anyway?"):
                    return False
        
        return True

    def refresh_rainmeter(self):
        if not self.ini_file:
            messagebox.showerror("Error", "No dock selected")
            self.set_status("Error: No dock selected")
            return
            
        try:
            dock_name = os.path.basename(self.ini_file).split('.')[0]
            
            dock_number = None
            if dock_name.startswith("LuckyDock") and dock_name[9:].isdigit():
                dock_number = dock_name[9:]
            
            if not dock_number:
                messagebox.showerror("Error", f"Invalid dock name format: {dock_name}")
                self.set_status(f"Error: Invalid dock name format: {dock_name}")
                return
            
            rainmeter_command = f'Rainmeter !Refresh "LuckyDock\\Dock{dock_number}\\{dock_name}"'
            
            subprocess.run(rainmeter_command, shell=True)
            self.set_status(f"Refreshed Rainmeter skin: LuckyDock\\Dock{dock_number}\\{dock_name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh Rainmeter skin: {str(e)}")
            self.set_status(f"Error refreshing Rainmeter skin: {str(e)}")

    def on_select_entry(self, event):
        selection = self.entries_listbox.curselection()
        if selection:
            index = selection[0]
            if self.selected_entry_index is not None:
                self.entries_listbox.itemconfig(self.selected_entry_index, bg=self.listbox_bg)
            
            self.selected_entry_index = index
            self.entries_listbox.itemconfig(index, bg=self.selected_bg)
            
            entry = self.entries[index]
            self.name_entry.delete(0, tk.END)
            self.name_entry.insert(0, entry["name"])
            
            self.app_path_entry.delete(0, tk.END)
            self.app_path_entry.insert(0, entry["app_path"])
            
            self.icon_path_entry.delete(0, tk.END)
            self.icon_path_entry.insert(0, entry["icon_path"])
            
            self.set_status(f"Selected entry: {entry['name']}")

    def update_entry(self):
        if self.selected_entry_index is None:
            messagebox.showerror("Error", "No entry selected")
            self.set_status("Error: No entry selected")
            return
        
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Validation Error", "Entry name cannot be empty")
            self.set_status("Error: Entry name cannot be empty")
            return
        
        is_separator = name.startswith("Separator_")
        
        self.entries[self.selected_entry_index] = {
            "name": name,
            "app_path": self.app_path_entry.get(),
            "icon_path": self.icon_path_entry.get(),
            "is_separator": is_separator
        }
        
        self.entries_listbox.delete(self.selected_entry_index)
        self.entries_listbox.insert(self.selected_entry_index, name)
        
        if is_separator:
            self.entries_listbox.itemconfig(self.selected_entry_index, fg=self.separator_color)
        else:
            self.entries_listbox.itemconfig(self.selected_entry_index, fg=self.fg_color)
            
        self.entries_listbox.selection_set(self.selected_entry_index)
        
        self.set_status(f"Updated entry: {name}")

    def add_entry(self):
        if not self.ini_file:
            messagebox.showerror("Error", "No dock selected")
            self.set_status("Error: No dock selected")
            return
            
        new_entry = {
            "name": "New Entry",
            "app_path": '""',
            "icon_path": '""',
            "is_separator": False
        }
        
        self.entries.append(new_entry)
        self.entries_listbox.insert(tk.END, new_entry["name"])
        self.entries_listbox.selection_clear(0, tk.END)
        self.entries_listbox.selection_set(tk.END)
        self.entries_listbox.see(tk.END)
        self.on_select_entry(None)
        
        self.set_status("Added new entry")
    
    def add_separator(self):
        if not self.ini_file:
            messagebox.showerror("Error", "No dock selected")
            self.set_status("Error: No dock selected")
            return
        
        separator_count = sum(1 for entry in self.entries if entry.get("is_separator", False))
        new_name = f"Separator_{separator_count + 1}"
        
        separator_icon = os.path.join(self.base_dir, "@Resources", "separator.png")
        
        new_entry = {
            "name": new_name,
            "app_path": '""',
            "icon_path": separator_icon,
            "is_separator": True
        }
        
        self.entries.append(new_entry)
        self.entries_listbox.insert(tk.END, new_name)
        self.entries_listbox.itemconfig(tk.END, fg=self.separator_color)
        self.entries_listbox.selection_clear(0, tk.END)
        self.entries_listbox.selection_set(tk.END)
        self.entries_listbox.see(tk.END)
        self.on_select_entry(None)
        
        self.set_status("Added new separator")

    def remove_entry(self):
        if self.selected_entry_index is None:
            messagebox.showerror("Error", "No entry selected")
            self.set_status("Error: No entry selected")
            return
        
        entry_name = self.entries[self.selected_entry_index]["name"]
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{entry_name}'?"):
            self.set_status("Delete cancelled")
            return
        
        index = self.selected_entry_index
        del self.entries[index]
        self.entries_listbox.delete(index)
        self.selected_entry_index = None
        
        if self.entries:
            new_index = min(index, len(self.entries) - 1)
            self.entries_listbox.selection_set(new_index)
            self.on_select_entry(None)
        else:
            self.name_entry.delete(0, tk.END)
            self.app_path_entry.delete(0, tk.END)
            self.icon_path_entry.delete(0, tk.END)
        
        self.set_status(f"Removed entry: {entry_name}")

    def move_entry(self, direction):
        if self.selected_entry_index is None:
            messagebox.showerror("Error", "No entry selected")
            self.set_status("Error: No entry selected")
            return
        
        index = self.selected_entry_index
        new_index = index + direction
        
        if 0 <= new_index < len(self.entries):
            self.entries[index], self.entries[new_index] = self.entries[new_index], self.entries[index]
            
            self.entries_listbox.delete(index)
            self.entries_listbox.insert(new_index, self.entries[new_index]["name"])
            
            if self.entries[new_index].get("is_separator", False):
                self.entries_listbox.itemconfig(new_index, fg=self.separator_color)
            else:
                self.entries_listbox.itemconfig(new_index, fg=self.fg_color)
                
            self.entries_listbox.selection_clear(0, tk.END)
            self.entries_listbox.selection_set(new_index)
            self.entries_listbox.see(new_index)
            self.selected_entry_index = new_index
            
            direction_text = "up" if direction < 0 else "down"
            self.set_status(f"Moved entry {direction_text}")

if __name__ == "__main__":
    dock_number = None
    if len(sys.argv) > 1:
        try:
            dock_arg = sys.argv[1]
            if dock_arg.isdigit():
                dock_number = dock_arg
            elif dock_arg.startswith("LuckyDock") and dock_arg[9:].isdigit():
                dock_number = dock_arg[9:]
        except Exception:
            pass
    
    root = tk.Tk()
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, "luckydock_icon.ico")
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass
        
    app = LuckyDockManager(root, dock_number)
    root.mainloop()