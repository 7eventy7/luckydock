import tkinter as tk
from tkinter import messagebox, font, filedialog, ttk, colorchooser
import configparser
import os
import subprocess
import re
import sys
import shutil
import time
from typing import Dict, List, Tuple, Optional, Any, Union

try:
    from PIL import Image, ImageTk
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

# Windows-specific console hiding
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


class Constants:
    """Centralized constants and configuration values."""
    
    # Application constants
    APP_TITLE = "LuckyDockManager"
    APP_VERSION = "1.0"
    
    # Window dimensions
    WINDOW_WIDTH = 1920
    WINDOW_HEIGHT = 1280
    
    # Color scheme
    BG_COLOR = "#1e1e1e"
    FG_COLOR = "#ffffff"
    ENTRY_BG = "#2d2d2d"
    BUTTON_BG = "#3e3e3e"
    BUTTON_HOVER_BG = "#4a4a4a"
    LISTBOX_BG = "#252525"
    SELECTED_BG = "#0078d7"
    STATUS_BG = "#1a1a1a"
    ACCENT_COLOR = "#007acc"
    SEPARATOR_COLOR = "#888888"
    SUCCESS_COLOR = "#28a745"
    DANGER_COLOR = "#dc3545"
    SELECTED_DOCK_BORDER = "#ffffff"
    
    # Default values
    DEFAULT_CORNER_RADIUS = 10
    DEFAULT_DOCK_SIZE = 60
    DEFAULT_BACKGROUND_COLOR = (0, 0, 0)
    DEFAULT_OPACITY = 175
    DEFAULT_ORIENTATION = "Vertical"
    DEFAULT_TOOLTIP_DELAY = 500
    DEFAULT_TOOLTIP_FONT = "Segoe UI"
    DEFAULT_TOOLTIP_FONT_SIZE = 9
    DEFAULT_BACKGROUND_HEIGHT = 80
    
    # Value ranges
    MIN_CORNER_RADIUS = 0
    MAX_CORNER_RADIUS = 50
    MIN_DOCK_SIZE = 40
    MAX_DOCK_SIZE = 70
    MIN_OPACITY = 0
    MAX_OPACITY = 255
    MIN_TOOLTIP_DELAY = 100
    MAX_TOOLTIP_DELAY = 2000
    
    # File and path constants
    TEMPLATE_FILENAME = "template.ini"
    RESOURCES_DIR = "@Resources"
    IMAGES_DIR = "images"
    LOGO_FILENAME = "logo.png"
    SEPARATOR_ICON = "separator.png"
    ICON_FILENAME = "luckydock_icon.ico"
    
    # Rainmeter constants
    RAINMETER_GROUP = "Taskbar"
    RAINMETER_UPDATE_RATE = 100
    CONFIG_PREFIX = "LuckyDock"
    DOCK_PREFIX = "Dock"
    
    # UI constants
    FONT_FAMILY = "Segoe UI"
    TITLE_FONT_SIZE = 36
    HEADER_FONT_SIZE = 16
    BUTTON_FONT_SIZE = 10
    ENTRY_FONT_SIZE = 12
    LISTBOX_FONT_SIZE = 16
    LOGO_MAX_HEIGHT = 80
    
    # Widget dimensions
    BUTTON_PADDING = 6
    FRAME_PADDING = 15
    WIDGET_PADDING = 10
    SCALE_LENGTH = 150
    ENTRY_WIDTH = 6
    COMBO_WIDTH = 20
    LISTBOX_SELECTION_WIDTH = 50
    UPDATE_BUTTON_WIDTH = 110
    
    # File type filters
    EXECUTABLE_FILTER = [("Executable files", "*.exe"), ("All files", "*.*")]
    IMAGE_FILTER = [
        ("Image files", "*.png *.jpg *.jpeg *.ico *.bmp *.gif"),
        ("All files", "*.*")
    ]
    
    # Excluded INI sections
    EXCLUDED_SECTIONS = {
        "Variables", "Rainmeter", "Metadata", "MeasureAppCount",
        "Style", "MeasureFlash", "BackgroundMeter"
    }
    
    # Font options
    FONT_OPTIONS = [
        "Segoe UI", "Arial", "Calibri", "Tahoma", "Verdana",
        "Times New Roman", "Georgia", "Comic Sans MS", "Impact"
    ]
    
    # Orientation options
    ORIENTATION_OPTIONS = ["Vertical", "Horizontal"]
    
    # Status messages
    STATUS_READY = "Ready"
    STATUS_LOADING = "Loading..."
    STATUS_SAVING = "Saving changes..."
    STATUS_CREATED = "Created new dock"
    STATUS_DELETED = "Successfully deleted dock"
    STATUS_UPDATED = "Updated entry"
    STATUS_MOVED = "Moved entry"
    STATUS_ADDED = "Added new entry"
    STATUS_REMOVED = "Removed entry"
    
    # Error messages
    ERROR_NO_DOCK = "No dock selected"
    ERROR_INVALID_DOCK = "Invalid dock format"
    ERROR_TEMPLATE_NOT_FOUND = "Template file not found"
    ERROR_DOCK_EXISTS = "Dock directory already exists"
    ERROR_NO_ENTRY = "No entry selected"
    ERROR_EMPTY_NAME = "Entry name cannot be empty"
    ERROR_DUPLICATE_NAME = "Duplicate entry name"
    ERROR_RAINMETER_NOT_FOUND = "Could not find Rainmeter.exe"
    
    # Timeouts and delays
    COMMAND_TIMEOUT = 10
    REFRESH_DELAY = 0.5
    UNLOAD_DELAY = 1.0


class DockEntry:
    """Represents a single dock entry with its properties."""
    
    def __init__(self, name: str, app_path: str = '""', icon_path: str = '""', is_separator: bool = False):
        self.name = name
        self.app_path = app_path
        self.icon_path = icon_path
        self.is_separator = is_separator
    
    def to_dict(self) -> Dict[str, Union[str, bool]]:
        """Convert entry to dictionary format."""
        return {
            "name": self.name,
            "app_path": self.app_path,
            "icon_path": self.icon_path,
            "is_separator": self.is_separator
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Union[str, bool]]) -> 'DockEntry':
        """Create entry from dictionary format."""
        return cls(
            name=data.get("name", ""),
            app_path=data.get("app_path", '""'),
            icon_path=data.get("icon_path", '""'),
            is_separator=data.get("is_separator", False)
        )


class DockValidator:
    """Handles validation logic for dock entries and configurations."""
    
    @staticmethod
    def validate_entries(entries: List[DockEntry], show_dialogs: bool = True) -> bool:
        """Validate all entries in the dock."""
        for i, entry in enumerate(entries):
            if not DockValidator.validate_single_entry(entry, i, entries, show_dialogs):
                return False
        return True
    
    @staticmethod
    def validate_single_entry(entry: DockEntry, index: int, all_entries: List[DockEntry], show_dialogs: bool = True) -> bool:
        """Validate a single entry."""
        # Check for empty name
        if not entry.name.strip():
            if show_dialogs:
                messagebox.showerror("Validation Error", f"Entry #{index+1} has an empty name")
            else:
                raise ValueError(f"Entry #{index+1} has an empty name")
            return False
        
        # Check for duplicate names
        name_count = sum(1 for e in all_entries if e.name == entry.name)
        if name_count > 1:
            if show_dialogs:
                messagebox.showerror("Validation Error", f"Duplicate entry name: {entry.name}")
            else:
                raise ValueError(f"Duplicate entry name: {entry.name}")
            return False
        
        # Check app path for non-separators
        if not entry.is_separator:
            if not entry.app_path.strip() or entry.app_path == '""':
                if show_dialogs:
                    return messagebox.askyesno("Confirm", 
                                             f"Entry '{entry.name}' has an empty app path. Save anyway?")
                # For silent validation, allow empty app paths
        
        # Check icon path
        if not entry.icon_path.strip() or entry.icon_path == '""':
            if show_dialogs:
                return messagebox.askyesno("Confirm", 
                                         f"Entry '{entry.name}' has an empty icon path. Save anyway?")
            # For silent validation, allow empty icon paths
        
        return True
    
    @staticmethod
    def validate_dock_name(dock_name: str) -> bool:
        """Validate dock name format."""
        return (dock_name.startswith(Constants.CONFIG_PREFIX) and 
                dock_name[len(Constants.CONFIG_PREFIX):].isdigit())
    
    @staticmethod
    def sanitize_section_name(name: str) -> str:
        """Convert display name to a valid Rainmeter section name."""
        # Remove or replace problematic characters
        sanitized = re.sub(r'[^a-zA-Z0-9_]', '_', name)
        # Remove multiple consecutive underscores
        sanitized = re.sub(r'_+', '_', sanitized)
        # Remove leading/trailing underscores
        sanitized = sanitized.strip('_')
        # Ensure it's not empty
        if not sanitized:
            sanitized = "Entry"
        return sanitized


class RainmeterInterface:
    """Handles all Rainmeter command execution and interactions."""
    
    def __init__(self):
        self._rainmeter_exe = self._find_rainmeter_executable()
    
    def _find_rainmeter_executable(self) -> Optional[str]:
        """Find the Rainmeter executable."""
        rainmeter_paths = [
            r"C:\Program Files\Rainmeter\Rainmeter.exe",
            r"C:\Program Files (x86)\Rainmeter\Rainmeter.exe",
            "Rainmeter.exe"
        ]
        
        for path in rainmeter_paths:
            if path == "Rainmeter.exe":
                try:
                    subprocess.run([path, "!About"], capture_output=True, 
                                 timeout=3, check=False)
                    return path
                except:
                    continue
            elif os.path.exists(path):
                return path
        return None
    
    def execute_command(self, command: str) -> bool:
        """Execute a Rainmeter command."""
        if not self._rainmeter_exe:
            raise RuntimeError(Constants.ERROR_RAINMETER_NOT_FOUND)
        
        try:
            if self._rainmeter_exe == "Rainmeter.exe":
                full_command = f'Rainmeter.exe {command}'
            else:
                full_command = f'"{self._rainmeter_exe}" {command}'
            
            result = subprocess.run(full_command, shell=True, capture_output=True, 
                                  text=True, timeout=Constants.COMMAND_TIMEOUT)
            return result.returncode == 0
        except Exception:
            return False
    
    def refresh_skin(self, dock_number: str) -> bool:
        """Refresh a dock skin."""
        config_name = f"{Constants.CONFIG_PREFIX}\\{Constants.DOCK_PREFIX}{dock_number}"
        skin_file = f"{Constants.CONFIG_PREFIX}{dock_number}.ini"
        
        # Try to activate first, then refresh
        activate_success = self.execute_command(f'!ActivateConfig "{config_name}" "{skin_file}"')
        time.sleep(Constants.REFRESH_DELAY)
        refresh_success = self.execute_command(f'!Refresh "{config_name}"')
        
        return activate_success or refresh_success
    
    def unload_skin(self, dock_number: str) -> bool:
        """Unload a dock skin."""
        config_name = f"{Constants.CONFIG_PREFIX}\\{Constants.DOCK_PREFIX}{dock_number}"
        
        commands_to_try = [
            f'!UnloadSkin "{config_name}"',
            f'!DeactivateConfig "{config_name}"',
            f'!UnloadSkin "{Constants.CONFIG_PREFIX}\\{Constants.DOCK_PREFIX}{dock_number}\\{Constants.CONFIG_PREFIX}{dock_number}.ini"'
        ]
        
        for command in commands_to_try:
            if self.execute_command(command):
                time.sleep(Constants.REFRESH_DELAY)
                return True
        
        return False


class DockConfigManager:
    """Manages dock configuration file operations."""
    
    def __init__(self, base_dir: str):
        self.base_dir = base_dir
        self.template_file = os.path.join(base_dir, Constants.RESOURCES_DIR, Constants.TEMPLATE_FILENAME)
    
    def find_available_docks(self) -> List[str]:
        """Find all available dock configurations."""
        docks = []
        for item in os.listdir(self.base_dir):
            if item.startswith(Constants.DOCK_PREFIX) and os.path.isdir(os.path.join(self.base_dir, item)):
                dock_dir = os.path.join(self.base_dir, item)
                for file in os.listdir(dock_dir):
                    if file.startswith(Constants.CONFIG_PREFIX) and file.endswith(".ini"):
                        dock_name = os.path.splitext(file)[0]
                        docks.append(dock_name)
        return sorted(docks)
    
    def get_next_dock_number(self, existing_docks: List[str]) -> int:
        """Find the next available dock number."""
        existing_numbers = []
        for dock in existing_docks:
            if dock.startswith(Constants.CONFIG_PREFIX) and dock[len(Constants.CONFIG_PREFIX):].isdigit():
                existing_numbers.append(int(dock[len(Constants.CONFIG_PREFIX):]))
        
        if not existing_numbers:
            return 1
        
        existing_numbers.sort()
        for i in range(1, max(existing_numbers) + 2):
            if i not in existing_numbers:
                return i
        return max(existing_numbers) + 1
    
    def load_entries_from_ini(self, ini_file: str) -> List[DockEntry]:
        """Load entries from an INI file."""
        if not os.path.exists(ini_file):
            raise FileNotFoundError(f"Could not find {ini_file}")
        
        entries = []
        config = configparser.ConfigParser()
        config.read(ini_file)
        
        for section in config.sections():
            if self._should_exclude_section(section):
                continue
            
            # Use DisplayName if it exists, otherwise fall back to section name
            name = config[section].get("DisplayName", section)
            app_path = self._extract_app_path(config[section])
            icon_path = config[section].get("ImageName", "")
            is_separator = name.startswith("Separator_") or section.startswith("Separator_")
            
            entry = DockEntry(name, app_path, icon_path, is_separator)
            entries.append(entry)
        
        return entries
    
    def _should_exclude_section(self, section: str) -> bool:
        """Check if a section should be excluded from entries."""
        return (section in Constants.EXCLUDED_SECTIONS or 
                section.endswith("_TooltipBG") or 
                section.endswith("_TooltipText") or
                section.startswith("999_Tooltip_"))
    
    def _extract_app_path(self, section: configparser.SectionProxy) -> str:
        """Extract app path from section actions."""
        for action_key in ["LeftMouseUpAction", "LeftMouseDoubleClickAction"]:
            action_string = section.get(action_key, "")
            if action_string:
                match = re.search(r'\["([^"]*)"\]', action_string)
                if match:
                    return f'"{match.group(1)}"'
        return ""
    
    def load_appearance_settings(self, ini_file: str) -> Dict[str, Any]:
        """Load appearance settings from INI file."""
        if not os.path.exists(ini_file):
            return self._get_default_appearance_settings()
        
        try:
            config = configparser.ConfigParser()
            config.read(ini_file)
            
            if 'Variables' not in config:
                return self._get_default_appearance_settings()
            
            variables = config['Variables']
            
            # Parse RGB color
            bg_color_str = variables.get('BackgroundColor', '0,0,0')
            try:
                rgb_parts = [int(x.strip()) for x in bg_color_str.split(',')]
                current_color = tuple(rgb_parts) if len(rgb_parts) == 3 else Constants.DEFAULT_BACKGROUND_COLOR
            except:
                current_color = Constants.DEFAULT_BACKGROUND_COLOR
            
            # Parse orientation
            orientation_value = variables.get('Orientation', '0')
            try:
                orientation = "Horizontal" if int(orientation_value) == 1 else "Vertical"
            except:
                orientation = Constants.DEFAULT_ORIENTATION
            
            return {
                'corner_radius': self._safe_int_parse(variables.get('CornerRadius'), Constants.DEFAULT_CORNER_RADIUS),
                'dock_size': self._safe_int_parse(variables.get('IconSize'), Constants.DEFAULT_DOCK_SIZE),
                'current_color': current_color,
                'opacity': self._safe_int_parse(variables.get('BackgroundOpacity'), Constants.DEFAULT_OPACITY),
                'orientation': orientation,
                'double_click': variables.get('DoubleClickMode', '0') == '1',
                'tooltips': variables.get('ShowTooltips', '0') == '1',
                'tooltip_delay': self._safe_int_parse(variables.get('TooltipDelay'), Constants.DEFAULT_TOOLTIP_DELAY),
                'tooltip_font': variables.get('TooltipFont', Constants.DEFAULT_TOOLTIP_FONT),
                'tooltip_font_size': self._safe_int_parse(variables.get('TooltipFontSize'), Constants.DEFAULT_TOOLTIP_FONT_SIZE)
            }
            
        except Exception:
            return self._get_default_appearance_settings()
    
    def _safe_int_parse(self, value: Optional[str], default: int) -> int:
        """Safely parse an integer value with fallback."""
        try:
            return int(value) if value else default
        except (ValueError, TypeError):
            return default
    
    def _get_default_appearance_settings(self) -> Dict[str, Any]:
        """Get default appearance settings."""
        return {
            'corner_radius': Constants.DEFAULT_CORNER_RADIUS,
            'dock_size': Constants.DEFAULT_DOCK_SIZE,
            'current_color': Constants.DEFAULT_BACKGROUND_COLOR,
            'opacity': Constants.DEFAULT_OPACITY,
            'orientation': Constants.DEFAULT_ORIENTATION,
            'double_click': False,
            'tooltips': False,
            'tooltip_delay': Constants.DEFAULT_TOOLTIP_DELAY,
            'tooltip_font': Constants.DEFAULT_TOOLTIP_FONT,
            'tooltip_font_size': Constants.DEFAULT_TOOLTIP_FONT_SIZE
        }
    
    def save_entries_to_ini(self, ini_file: str, entries: List[DockEntry], appearance_settings: Dict[str, Any]) -> None:
        """Save entries and settings to INI file."""
        config = configparser.ConfigParser()
        config.read(ini_file)
        
        # Remove old entry sections
        self._remove_entry_sections(config)
        
        # Add entries
        for i, entry in enumerate(entries):
            self._add_entry_to_config(config, entry, i, appearance_settings)
        
        # Update variables
        self._update_config_variables(config, entries, appearance_settings)
        
        # Update background meter
        self._update_background_meter(config)
        
        # Write file
        with open(ini_file, 'w') as configfile:
            config.write(configfile)
    
    def _remove_entry_sections(self, config: configparser.ConfigParser) -> None:
        """Remove old entry sections from config."""
        sections_to_remove = [
            section for section in config.sections() 
            if section not in Constants.EXCLUDED_SECTIONS
        ]
        
        for section in sections_to_remove:
            config.remove_section(section)
    
    def _add_entry_to_config(self, config: configparser.ConfigParser, entry: DockEntry, 
                           index: int, appearance_settings: Dict[str, Any]) -> None:
        """Add a single entry to the config."""
        section_name = DockValidator.sanitize_section_name(entry.name)
        
        # Ensure unique section names
        original_section = section_name
        counter = 1
        while section_name in config.sections():
            section_name = f"{original_section}_{counter}"
            counter += 1
        
        config[section_name] = {}
        config[section_name]["Meter"] = "Image"
        config[section_name]["MeterStyle"] = "Style"
        config[section_name]["ImageName"] = entry.icon_path
        config[section_name]["W"] = "#IconSize#"
        config[section_name]["H"] = "#IconSize#"
        config[section_name]["DisplayName"] = entry.name
        
        # Position calculation based on orientation
        is_horizontal = appearance_settings['orientation'] == "Horizontal"
        if is_horizontal:
            config[section_name]["X"] = f"(#PadLeft# + {index} * (#IconSize# + #HorizontalGap#))"
            config[section_name]["Y"] = "((#BackgroundHeight# - #IconSize#) / 2)"
        else:
            config[section_name]["X"] = "((#BackgroundWidth# - #IconSize#) / 2)"
            config[section_name]["Y"] = f"(#PadTop# + {index} * (#IconSize# + #VerticalGap#))"
        
        if entry.is_separator:
            config[section_name]["DynamicVariables"] = "1"
        else:
            self._add_mouse_actions(config, section_name, entry, appearance_settings['double_click'])
            self._add_hover_actions(config, section_name, entry.name, index, 
                                   appearance_settings['tooltips'], is_horizontal)
            config[section_name]["DynamicVariables"] = "1"
    
    def _add_mouse_actions(self, config: configparser.ConfigParser, section_name: str, 
                          entry: DockEntry, double_click_enabled: bool) -> None:
        """Add mouse actions to entry section."""
        mouse_action = f'[{entry.app_path}][!SetVariable CurrentIcon "{section_name}"][!CommandMeasure MeasureFlash "Execute 1"]'
        
        if double_click_enabled:
            config[section_name]["LeftMouseDoubleClickAction"] = mouse_action
        else:
            config[section_name]["LeftMouseUpAction"] = mouse_action
    
    def _add_hover_actions(self, config: configparser.ConfigParser, section_name: str, 
                          display_name: str, index: int, tooltips_enabled: bool, is_horizontal: bool) -> None:
        """Add hover actions with tooltips."""
        # Base hover actions
        base_hover_in = (f'[!SetOption {section_name} W "(#IconSize# + #HoverSize#)"]'
                        f'[!SetOption {section_name} H "(#IconSize# + #HoverSize#)"]'
                        f'[!SetOption {section_name} X "([{section_name}:X] - (#HoverSize# / 2))"]'
                        f'[!SetOption {section_name} Y "([{section_name}:Y] - (#HoverSize# / 2))"]'
                        f'[!UpdateMeter {section_name}][!Redraw]')
        
        base_hover_out = (f'[!SetOption {section_name} W "#IconSize#"]'
                         f'[!SetOption {section_name} H "#IconSize#"]'
                         f'[!SetOption {section_name} X "(#Horizontal# ? (#PadLeft# + {index} * (#IconSize# + #HorizontalGap#)) : ((#BackgroundWidth# - #IconSize#) / 2))"]'
                         f'[!SetOption {section_name} Y "(#Vertical# ? (#PadTop# + {index} * (#IconSize# + #VerticalGap#)) : ((#BackgroundHeight# - #IconSize#) / 2))"]'
                         f'[!UpdateMeter {section_name}][!Redraw]')
        
        if tooltips_enabled:
            tooltip_bg_name = f"999_Tooltip_{index}_BG"
            tooltip_text_name = f"999_Tooltip_{index}_Text"
            
            self._create_tooltip_sections(config, tooltip_bg_name, tooltip_text_name, 
                                        display_name, index, is_horizontal)
            
            tooltip_show = (f'[!SetOption {tooltip_bg_name} Hidden "0"]'
                           f'[!SetOption {tooltip_text_name} Hidden "0"]'
                           f'[!UpdateMeter {tooltip_bg_name}]'
                           f'[!UpdateMeter {tooltip_text_name}][!Redraw]')
            
            tooltip_hide = (f'[!SetOption {tooltip_bg_name} Hidden "1"]'
                           f'[!SetOption {tooltip_text_name} Hidden "1"]'
                           f'[!UpdateMeter {tooltip_bg_name}]'
                           f'[!UpdateMeter {tooltip_text_name}][!Redraw]')
            
            config[section_name]["MouseOverAction"] = base_hover_in + tooltip_show
            config[section_name]["MouseLeaveAction"] = base_hover_out + tooltip_hide
        else:
            config[section_name]["MouseOverAction"] = base_hover_in
            config[section_name]["MouseLeaveAction"] = base_hover_out
    
    def _create_tooltip_sections(self, config: configparser.ConfigParser, tooltip_bg_name: str,
                               tooltip_text_name: str, display_name: str, index: int, is_horizontal: bool) -> None:
        """Create tooltip background and text sections."""
        if is_horizontal:
            tooltip_x = f"(#PadLeft# + {index} * (#IconSize# + #HorizontalGap#) - 10)"
            tooltip_y = f"(((#BackgroundHeight# - #IconSize#) / 2) - (#TooltipFontSize# + 20) + 40)"
            text_x = f"(#PadLeft# + {index} * (#IconSize# + #HorizontalGap#) + (#IconSize# / 2))"
            text_y = f"(((#BackgroundHeight# - #IconSize#) / 2) - (#TooltipFontSize# + 16) + 40)"
            
            config[tooltip_bg_name] = {
                "Meter": "Shape",
                "Shape": "Rectangle 0,0,(#IconSize# + 20),(#TooltipFontSize# + 16),5 | Fill Color #BackgroundColor#,200 | StrokeWidth 0",
                "X": tooltip_x,
                "Y": tooltip_y,
                "Hidden": "1",
                "DynamicVariables": "1"
            }
        else:
            tooltip_y = f"(#PadTop# + {index} * (#IconSize# + #VerticalGap#) + (#IconSize# / 2) - (#TooltipFontSize# / 2) - 4)"
            text_x = "(#BackgroundWidth# / 2)"
            text_y = f"(#PadTop# + {index} * (#IconSize# + #VerticalGap#) + (#IconSize# / 2) - (#TooltipFontSize# / 2))"
            
            config[tooltip_bg_name] = {
                "Meter": "Shape",
                "Shape": "Rectangle 0,0,(#BackgroundWidth# - 4),(#TooltipFontSize# + 16),5 | Fill Color #BackgroundColor#,200 | StrokeWidth 0",
                "X": "2",
                "Y": tooltip_y,
                "Hidden": "1",
                "DynamicVariables": "1"
            }
            
            text_x = "(#BackgroundWidth# / 2)"
        
        config[tooltip_text_name] = {
            "Meter": "String",
            "Text": display_name,
            "FontColor": "255,255,255,255",
            "FontSize": "#TooltipFontSize#",
            "FontFace": "#TooltipFont#",
            "StringAlign": "Center",
            "X": text_x,
            "Y": text_y,
            "Hidden": "1",
            "DynamicVariables": "1"
        }
    
    def _update_config_variables(self, config: configparser.ConfigParser, entries: List[DockEntry], 
                               appearance_settings: Dict[str, Any]) -> None:
        """Update configuration variables."""
        if 'Variables' not in config:
            config.add_section('Variables')
        
        config.set('Variables', 'AppCount', str(len(entries)))
        config.set('Variables', 'CornerRadius', str(appearance_settings['corner_radius']))
        config.set('Variables', 'IconSize', str(appearance_settings['dock_size']))
        
        rgb = f"{appearance_settings['current_color'][0]},{appearance_settings['current_color'][1]},{appearance_settings['current_color'][2]}"
        config.set('Variables', 'BackgroundColor', rgb)
        config.set('Variables', 'BackgroundOpacity', str(appearance_settings['opacity']))
        
        # Orientation settings
        is_horizontal = appearance_settings['orientation'] == "Horizontal"
        config.set('Variables', 'Orientation', '1' if is_horizontal else '0')
        config.set('Variables', 'Horizontal', '1' if is_horizontal else '0')
        config.set('Variables', 'Vertical', '0' if is_horizontal else '1')
        
        # Ensure BackgroundHeight exists
        if not config.has_option('Variables', 'BackgroundHeight'):
            config.set('Variables', 'BackgroundHeight', str(Constants.DEFAULT_BACKGROUND_HEIGHT))
        
        # Behavior settings
        config.set('Variables', 'DoubleClickMode', '1' if appearance_settings['double_click'] else '0')
        config.set('Variables', 'ShowTooltips', '1' if appearance_settings['tooltips'] else '0')
        config.set('Variables', 'TooltipDelay', str(appearance_settings['tooltip_delay']))
        config.set('Variables', 'TooltipFont', appearance_settings['tooltip_font'])
        config.set('Variables', 'TooltipFontSize', str(appearance_settings['tooltip_font_size']))
    
    def _update_background_meter(self, config: configparser.ConfigParser) -> None:
        """Update background meter configuration."""
        if 'BackgroundMeter' not in config:
            config.add_section('BackgroundMeter')
        
        config.set('BackgroundMeter', 'Meter', 'Shape')
        config.set('BackgroundMeter', 'Shape', 
                  'Rectangle 0,0,(#Horizontal# ? ((#IconSize# + #HorizontalGap#) * #AppCount# + #PadLeft# + #PadRight# - #HorizontalGap#) : #BackgroundWidth#),(#Vertical# ? ((#IconSize# + #VerticalGap#) * #AppCount# + #PadTop# + #PadBottom# - #VerticalGap#) : #BackgroundHeight#),#CornerRadius# | Fill Color #BackgroundColor#,#BackgroundOpacity# | StrokeWidth 0')
        config.set('BackgroundMeter', 'DynamicVariables', '1')
    
    def create_new_dock(self, dock_number: int) -> str:
        """Create a new dock from template."""
        if not os.path.exists(self.template_file):
            raise FileNotFoundError(Constants.ERROR_TEMPLATE_NOT_FOUND)
        
        dock_name = f"Lucky Dock {dock_number}"
        dock_folder_name = f"{Constants.DOCK_PREFIX}{dock_number}"
        dock_file_name = f"{Constants.CONFIG_PREFIX}{dock_number}"
        dock_dir = os.path.join(self.base_dir, dock_folder_name)
        
        if os.path.exists(dock_dir):
            raise FileExistsError(Constants.ERROR_DOCK_EXISTS)
        
        try:
            os.makedirs(dock_dir)
            new_dock_file = os.path.join(dock_dir, f"{dock_file_name}.ini")
            shutil.copy2(self.template_file, new_dock_file)
            self._update_dock_template(new_dock_file, dock_number, dock_name)
            return dock_file_name
        except Exception as e:
            if os.path.exists(dock_dir):
                try:
                    shutil.rmtree(dock_dir)
                except:
                    pass
            raise e
    
    def _update_dock_template(self, file_path: str, dock_number: int, dock_name: str) -> None:
        """Update template file with correct dock number and name."""
        # Update context action and name
        with open(file_path, 'r') as file:
            lines = file.readlines()
        
        for i, line in enumerate(lines):
            if 'contextaction =' in line and 'LuckyDockManager.pyw' in line:
                lines[i] = re.sub(r'"\d+"(?=\s*])', f'"{dock_number}"', line)
            elif line.strip().startswith('name = "Lucky Dock'):
                lines[i] = f'name = "{dock_name}"\n'
        
        with open(file_path, 'w') as file:
            file.writelines(lines)
        
        # Set default appearance values
        self._set_template_defaults(file_path)
    
    def _set_template_defaults(self, file_path: str) -> None:
        """Set default values in template file."""
        try:
            config = configparser.ConfigParser()
            config.read(file_path)
            
            if 'Variables' in config:
                default_vars = {
                    'CornerRadius': str(Constants.DEFAULT_CORNER_RADIUS),
                    'IconSize': str(Constants.DEFAULT_DOCK_SIZE),
                    'BackgroundColor': f"{Constants.DEFAULT_BACKGROUND_COLOR[0]},{Constants.DEFAULT_BACKGROUND_COLOR[1]},{Constants.DEFAULT_BACKGROUND_COLOR[2]}",
                    'BackgroundOpacity': str(Constants.DEFAULT_OPACITY),
                    'Orientation': '0',
                    'Vertical': '1',
                    'Horizontal': '0',
                    'BackgroundHeight': str(Constants.DEFAULT_BACKGROUND_HEIGHT),
                    'DoubleClickMode': '0',
                    'ShowTooltips': '0',
                    'TooltipDelay': str(Constants.DEFAULT_TOOLTIP_DELAY),
                    'TooltipFont': Constants.DEFAULT_TOOLTIP_FONT,
                    'TooltipFontSize': str(Constants.DEFAULT_TOOLTIP_FONT_SIZE)
                }
                
                for var, value in default_vars.items():
                    config.set('Variables', var, value)
                
                # Update New_Entry section
                if 'New_Entry' in config:
                    config.set('New_Entry', 'DisplayName', 'New Entry')
                
                with open(file_path, 'w') as configfile:
                    config.write(configfile)
        except Exception:
            pass  # Continue if setting defaults fails
    
    def delete_dock(self, dock_number: str) -> None:
        """Delete a dock directory and files."""
        dock_dir = os.path.join(self.base_dir, f"{Constants.DOCK_PREFIX}{dock_number}")
        
        if os.path.exists(dock_dir):
            time.sleep(Constants.UNLOAD_DELAY)  # Allow Rainmeter to release handles
            shutil.rmtree(dock_dir)


class LuckyDockManager:
    """Main application class for the LuckyDock Manager."""
    
    def __init__(self, master: tk.Tk, dock_number: Optional[str] = None):
        self.master = master
        self._initialize_window()
        self._initialize_paths()
        self._initialize_managers()
        self._initialize_variables()
        self._initialize_dock_selection(dock_number)
        
        self.create_widgets()
        self.setup_keyboard_shortcuts()
        
        if self.available_docks:
            self.load_dock(self.current_dock.get())
    
    def _initialize_window(self) -> None:
        """Initialize the main window."""
        self.master.title(Constants.APP_TITLE)
        self.master.geometry(f"{Constants.WINDOW_WIDTH}x{Constants.WINDOW_HEIGHT}")
        self.master.configure(bg=Constants.BG_COLOR)
    
    def _initialize_paths(self) -> None:
        """Initialize file paths."""
        script_dir = os.path.dirname(os.path.abspath(__file__))
        self.base_dir = os.path.dirname(script_dir)
    
    def _initialize_managers(self) -> None:
        """Initialize manager classes."""
        self.config_manager = DockConfigManager(self.base_dir)
        self.rainmeter = RainmeterInterface()
        
        self.available_docks = self.config_manager.find_available_docks()
    
    def _initialize_variables(self) -> None:
        """Initialize GUI variables."""
        # Dock management
        self.current_dock = tk.StringVar()
        self.ini_file = None
        self.entries = []
        self.selected_entry_index = None
        self.status_message = tk.StringVar(value=Constants.STATUS_READY)
        
        # Appearance variables
        self.corner_radius_var = tk.IntVar(value=Constants.DEFAULT_CORNER_RADIUS)
        self.dock_size_var = tk.IntVar(value=Constants.DEFAULT_DOCK_SIZE)
        self.current_color = Constants.DEFAULT_BACKGROUND_COLOR
        self.opacity_var = tk.IntVar(value=Constants.DEFAULT_OPACITY)
        
        # Behavior variables
        self.double_click_var = tk.BooleanVar(value=False)
        self.tooltips_var = tk.BooleanVar(value=False)
        self.tooltip_delay_var = tk.IntVar(value=Constants.DEFAULT_TOOLTIP_DELAY)
        self.tooltip_font_var = tk.StringVar(value=Constants.DEFAULT_TOOLTIP_FONT)
        self.tooltip_font_size_var = tk.IntVar(value=Constants.DEFAULT_TOOLTIP_FONT_SIZE)
        self.orientation_var = tk.StringVar(value=Constants.DEFAULT_ORIENTATION)
    
    def _initialize_dock_selection(self, dock_number: Optional[str]) -> None:
        """Initialize dock selection based on available docks and provided dock number."""
        if self.available_docks:
            if dock_number is not None:
                dock_name = f"{Constants.CONFIG_PREFIX}{dock_number}"
                if dock_name in self.available_docks:
                    self.current_dock.set(dock_name)
                else:
                    self.current_dock.set(self.available_docks[0])
            else:
                self.current_dock.set(self.available_docks[0])
    
    def create_widgets(self) -> None:
        """Create and configure all GUI widgets."""
        self._configure_styles()
        
        main_container = ttk.Frame(self.master, style='TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self._create_header(main_container)
        self._create_dock_management_section(main_container)
        self._create_content_section(main_container)
        self._create_bottom_section(main_container)
        self._create_status_bar()
    
    def _configure_styles(self) -> None:
        """Configure ttk styles for dark theme."""
        self.style = ttk.Style()
        self.style.theme_use('default')
        
        # Basic styles
        style_configs = {
            'TFrame': {'background': Constants.BG_COLOR},
            'TLabel': {
                'background': Constants.BG_COLOR, 
                'foreground': Constants.FG_COLOR, 
                'font': (Constants.FONT_FAMILY, Constants.BUTTON_FONT_SIZE)
            },
            'TButton': {
                'background': Constants.BUTTON_BG, 
                'foreground': Constants.FG_COLOR,
                'font': (Constants.FONT_FAMILY, Constants.BUTTON_FONT_SIZE),
                'borderwidth': 0,
                'padding': Constants.BUTTON_PADDING,
                'relief': tk.FLAT
            },
            'TEntry': {
                'fieldbackground': Constants.ENTRY_BG, 
                'foreground': Constants.FG_COLOR,
                'borderwidth': 1,
                'relief': tk.SOLID
            },
            'TCombobox': {
                'fieldbackground': Constants.ENTRY_BG, 
                'foreground': Constants.FG_COLOR,
                'background': Constants.BUTTON_BG,
                'arrowcolor': Constants.FG_COLOR,
                'borderwidth': 1,
                'relief': tk.SOLID
            },
            'TScale': {
                'background': Constants.BG_COLOR,
                'troughcolor': Constants.ENTRY_BG,
                'borderwidth': 0,
                'lightcolor': Constants.ACCENT_COLOR,
                'darkcolor': Constants.ACCENT_COLOR
            },
            'TCheckbutton': {
                'background': Constants.BG_COLOR,
                'foreground': Constants.FG_COLOR,
                'font': (Constants.FONT_FAMILY, Constants.BUTTON_FONT_SIZE),
                'focuscolor': 'none'
            }
        }
        
        for style_name, config in style_configs.items():
            self.style.configure(style_name, **config)
        
        # Configure hover states
        self.style.map('TButton', background=[('active', Constants.BUTTON_HOVER_BG)])
        self.style.map('TCheckbutton', background=[('active', Constants.BG_COLOR)])
        
        # Configure combobox states
        self._configure_combobox_styles()
        
        # Special button styles
        self._configure_special_button_styles()
    
    def _configure_combobox_styles(self) -> None:
        """Configure combobox-specific styles."""
        self.style.map('TCombobox',
                      fieldbackground=[('readonly', Constants.ENTRY_BG),
                                     ('focus', Constants.ENTRY_BG)],
                      selectbackground=[('readonly', Constants.SELECTED_BG)],
                      selectforeground=[('readonly', Constants.FG_COLOR)],
                      background=[('active', Constants.BUTTON_HOVER_BG),
                                ('pressed', Constants.BUTTON_BG),
                                ('readonly', Constants.BUTTON_BG)],
                      arrowcolor=[('active', Constants.FG_COLOR)],
                      foreground=[('readonly', Constants.FG_COLOR)])
        
        # Configure dropdown list
        self.master.option_add('*TCombobox*Listbox.Background', Constants.LISTBOX_BG)
        self.master.option_add('*TCombobox*Listbox.Foreground', Constants.FG_COLOR)
        self.master.option_add('*TCombobox*Listbox.selectBackground', Constants.SELECTED_BG)
        self.master.option_add('*TCombobox*Listbox.selectForeground', Constants.FG_COLOR)
    
    def _configure_special_button_styles(self) -> None:
        """Configure special button styles (accent, success, danger)."""
        special_styles = {
            'Accent.TButton': (Constants.ACCENT_COLOR, '#0069b3'),
            'Success.TButton': (Constants.SUCCESS_COLOR, '#218838'),
            'Danger.TButton': (Constants.DANGER_COLOR, '#c82333')
        }
        
        for style_name, (bg_color, hover_color) in special_styles.items():
            self.style.configure(style_name, 
                                background=bg_color, 
                                foreground=Constants.FG_COLOR,
                                font=(Constants.FONT_FAMILY, Constants.BUTTON_FONT_SIZE, 'bold'),
                                padding=Constants.BUTTON_PADDING)
            self.style.map(style_name, background=[('active', hover_color)])
    
    def _create_header(self, parent: ttk.Frame) -> None:
        """Create header section with title/logo."""
        title_frame = ttk.Frame(parent, style='TFrame')
        title_frame.pack(pady=Constants.WIDGET_PADDING, fill=tk.X)
        
        title_container = ttk.Frame(title_frame, style='TFrame')
        title_container.pack(pady=5)
        
        self._load_logo(title_container)
        ttk.Separator(parent, orient='horizontal').pack(fill=tk.X, pady=Constants.WIDGET_PADDING)
    
    def _load_logo(self, parent_container: ttk.Frame) -> None:
        """Load and display the logo image, fallback to text if it fails."""
        logo_loaded = False
        
        if HAS_PIL:
            try:
                logo_path = os.path.join(self.base_dir, Constants.RESOURCES_DIR, 
                                       Constants.IMAGES_DIR, Constants.LOGO_FILENAME)
                if os.path.exists(logo_path):
                    logo_image = Image.open(logo_path)
                    if logo_image.height > Constants.LOGO_MAX_HEIGHT:
                        ratio = Constants.LOGO_MAX_HEIGHT / logo_image.height
                        new_width = int(logo_image.width * ratio)
                        logo_image = logo_image.resize((new_width, Constants.LOGO_MAX_HEIGHT), 
                                                     Image.Resampling.LANCZOS)
                    
                    self.logo_photo = ImageTk.PhotoImage(logo_image)
                    logo_label = ttk.Label(parent_container, image=self.logo_photo, style='TLabel')
                    logo_label.pack()
                    logo_loaded = True
            except Exception:
                pass
        
        if not logo_loaded:
            title_font = self._get_font_with_fallback(Constants.TITLE_FONT_SIZE, "bold", 
                                                    ["Impact", "Verdana"])
            title_label = ttk.Label(parent_container, text=Constants.APP_TITLE, 
                                   font=title_font, foreground="#00aaff", style='TLabel')
            title_label.pack()
    
    def _get_font_with_fallback(self, size: int, weight: str = "normal", 
                               families: Optional[List[str]] = None) -> font.Font:
        """Get font with fallback options."""
        if families is None:
            families = [Constants.FONT_FAMILY]
        
        for family in families:
            try:
                return font.Font(family=family, size=size, weight=weight)
            except tk.TclError:
                continue
        
        return font.Font(size=size, weight=weight)
    
    def _create_dock_management_section(self, parent: ttk.Frame) -> None:
        """Create dock management and appearance section."""
        dock_management_frame = ttk.Frame(parent, style='TFrame')
        dock_management_frame.pack(pady=Constants.WIDGET_PADDING, fill=tk.X)
        
        dock_controls_frame = ttk.Frame(dock_management_frame, style='TFrame')
        dock_controls_frame.pack(fill=tk.X, pady=Constants.WIDGET_PADDING)
        
        self._create_dock_selector(dock_controls_frame)
        self._create_appearance_controls(dock_controls_frame)
        
        ttk.Separator(parent, orient='horizontal').pack(fill=tk.X, pady=15)
    
    def _create_dock_selector(self, parent: ttk.Frame) -> None:
        """Create dock selector and management buttons."""
        dock_selector_container = ttk.Frame(parent, style='TFrame')
        dock_selector_container.pack(side=tk.LEFT, padx=(0, 20))
        
        # Dock selector box
        dock_selector_frame = ttk.Frame(dock_selector_container, style='TFrame', 
                                       padding=(Constants.FRAME_PADDING, 12))
        dock_selector_frame.configure(relief=tk.SOLID, borderwidth=1)
        dock_selector_frame.pack(pady=(0, Constants.WIDGET_PADDING))
        
        ttk.Label(dock_selector_frame, text="Selected Dock:", 
                 font=self._get_font_with_fallback(15, "bold"), style='TLabel').pack(pady=(0, Constants.WIDGET_PADDING))
        
        self.dock_combo = ttk.Combobox(dock_selector_frame, 
                                      textvariable=self.current_dock, 
                                      values=self.available_docks, 
                                      state="readonly", 
                                      font=self._get_font_with_fallback(Constants.HEADER_FONT_SIZE, "bold"),
                                      width=Constants.COMBO_WIDTH,
                                      style='TCombobox')
        self.dock_combo.pack()
        self.dock_combo.bind("<<ComboboxSelected>>", self.on_dock_selected)
        
        # Management buttons
        dock_buttons_frame = ttk.Frame(dock_selector_container, style='TFrame')
        dock_buttons_frame.pack(fill=tk.X)
        
        ttk.Button(dock_buttons_frame, text="Create New Dock", 
                  command=self.create_new_dock, style='Success.TButton').pack(fill=tk.X, pady=(0, 5))
        ttk.Button(dock_buttons_frame, text="Delete Dock", 
                  command=self.delete_dock, style='Danger.TButton').pack(fill=tk.X)
    
    def _create_appearance_controls(self, parent: ttk.Frame) -> None:
        """Create appearance and behavior controls."""
        appearance_frame = ttk.Frame(parent, style='TFrame', 
                                   padding=(Constants.FRAME_PADDING, 12))
        appearance_frame.configure(relief=tk.SOLID, borderwidth=1)
        appearance_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self._create_dock_options_control(appearance_frame)
        self._create_color_controls(appearance_frame)
        self._create_behavior_controls(appearance_frame)
    
    def _create_dock_options_control(self, parent: ttk.Frame) -> None:
        """Create dock options (corners and size) controls."""
        options_frame = ttk.Frame(parent, style='TFrame')
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(options_frame, text="Dock Options:", 
                 font=self._get_font_with_fallback(12, "bold"), style='TLabel').pack(anchor=tk.W, pady=(0, 5))
        
        # Corner radius control
        self._create_scale_control(options_frame, "Corner Radius:", self.corner_radius_var,
                                 Constants.MIN_CORNER_RADIUS, Constants.MAX_CORNER_RADIUS,
                                 self.on_radius_scale_change, self.on_radius_entry_change)
        
        # Size control
        self._create_scale_control(options_frame, "Icon Size:", self.dock_size_var,
                                 Constants.MIN_DOCK_SIZE, Constants.MAX_DOCK_SIZE,
                                 self.on_size_scale_change, self.on_size_entry_change)
    
    def _create_scale_control(self, parent: ttk.Frame, label: str, variable: tk.IntVar,
                            min_val: int, max_val: int, scale_callback, entry_callback) -> None:
        """Create a scale control with entry field."""
        controls = ttk.Frame(parent, style='TFrame')
        controls.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(controls, text=label, font=self._get_font_with_fallback(Constants.BUTTON_FONT_SIZE), 
                 style='TLabel').pack(side=tk.LEFT, padx=(0, Constants.WIDGET_PADDING))
        
        scale = ttk.Scale(controls, from_=min_val, to=max_val, orient=tk.HORIZONTAL, 
                         variable=variable, length=Constants.SCALE_LENGTH,
                         command=scale_callback)
        scale.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, Constants.WIDGET_PADDING))
        
        entry = ttk.Entry(controls, textvariable=variable, 
                         width=Constants.ENTRY_WIDTH, font=self._get_font_with_fallback(Constants.BUTTON_FONT_SIZE))
        entry.pack(side=tk.RIGHT)
        entry.bind('<Return>', entry_callback)
        entry.bind('<FocusOut>', entry_callback)
    
    def _create_color_controls(self, parent: ttk.Frame) -> None:
        """Create color picker and opacity controls."""
        color_frame = ttk.Frame(parent, style='TFrame')
        color_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(color_frame, text="Dock Color:", 
                 font=self._get_font_with_fallback(12, "bold"), style='TLabel').pack(anchor=tk.W, pady=(0, 5))
        
        color_controls = ttk.Frame(color_frame, style='TFrame')
        color_controls.pack(fill=tk.X)
        
        # Color picker
        self._create_color_picker(color_controls)
        
        # Opacity control
        self._create_scale_control(color_frame, "Opacity:", self.opacity_var,
                                 Constants.MIN_OPACITY, Constants.MAX_OPACITY,
                                 lambda v: None, lambda e: None)
    
    def _create_color_picker(self, parent: ttk.Frame) -> None:
        """Create color picker section."""
        color_picker_frame = ttk.Frame(parent, style='TFrame')
        color_picker_frame.pack(fill=tk.X, pady=(0, 8))
        
        self.color_preview = tk.Label(color_picker_frame, width=4, height=2, 
                                     bg="#000000", relief=tk.RAISED, borderwidth=2)
        self.color_preview.pack(side=tk.LEFT, padx=(0, Constants.WIDGET_PADDING))
        
        ttk.Button(color_picker_frame, text="Choose Color", 
                  command=self.choose_color, width=15).pack(side=tk.LEFT, padx=(0, Constants.WIDGET_PADDING))
        
        # RGB entry
        manual_color_frame = ttk.Frame(color_picker_frame, style='TFrame')
        manual_color_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Label(manual_color_frame, text="RGB:", font=self._get_font_with_fallback(9), 
                 style='TLabel').pack(side=tk.LEFT, padx=(0, 5))
        self.rgb_entry = ttk.Entry(manual_color_frame, width=12, font=self._get_font_with_fallback(9))
        self.rgb_entry.pack(side=tk.LEFT)
        self.rgb_entry.bind('<Return>', self.on_manual_color_change)
        self.rgb_entry.bind('<FocusOut>', self.on_manual_color_change)
    
    def _create_behavior_controls(self, parent: ttk.Frame) -> None:
        """Create dock behavior controls."""
        behavior_label = ttk.Label(parent, text="Dock Behavior:", 
                                  font=self._get_font_with_fallback(12, "bold"), style='TLabel')
        behavior_label.pack(anchor=tk.W, pady=(15, Constants.WIDGET_PADDING))
        
        # Orientation control
        self._create_orientation_control(parent)
        
        # Double-click mode
        self._create_checkbox_control(parent, "Double-Click to Launch Apps", self.double_click_var)
        
        # Tooltips with font selection
        self._create_tooltip_controls(parent)
    
    def _create_orientation_control(self, parent: ttk.Frame) -> None:
        """Create orientation selection control."""
        orientation_frame = ttk.Frame(parent, style='TFrame')
        orientation_frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Label(orientation_frame, text="Orientation:", font=self._get_font_with_fallback(Constants.BUTTON_FONT_SIZE), 
                 style='TLabel').pack(side=tk.LEFT, padx=(0, Constants.WIDGET_PADDING))
        orientation_combo = ttk.Combobox(orientation_frame, textvariable=self.orientation_var,
                                        values=Constants.ORIENTATION_OPTIONS,
                                        state="readonly", width=12, 
                                        font=self._get_font_with_fallback(Constants.BUTTON_FONT_SIZE),
                                        style='TCombobox')
        orientation_combo.pack(side=tk.LEFT)
        orientation_combo.bind("<<ComboboxSelected>>", self.on_orientation_change)
    
    def _create_checkbox_control(self, parent: ttk.Frame, text: str, variable: tk.BooleanVar) -> None:
        """Create a checkbox control."""
        frame = ttk.Frame(parent, style='TFrame')
        frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Checkbutton(frame, text=text, variable=variable, style='TCheckbutton').pack(side=tk.LEFT)
    
    def _create_tooltip_controls(self, parent: ttk.Frame) -> None:
        """Create tooltip configuration controls."""
        tooltips_frame = ttk.Frame(parent, style='TFrame')
        tooltips_frame.pack(fill=tk.X, pady=(0, 8))
        
        ttk.Checkbutton(tooltips_frame, text="App Tooltips", 
                       variable=self.tooltips_var, style='TCheckbutton').pack(side=tk.LEFT, padx=(0, 15))
        
        # Tooltip font selection
        ttk.Label(tooltips_frame, text="Font:", font=self._get_font_with_fallback(9), 
                 style='TLabel').pack(side=tk.LEFT, padx=(0, 5))
        tooltip_font_combo = ttk.Combobox(tooltips_frame, textvariable=self.tooltip_font_var,
                                        values=Constants.FONT_OPTIONS,
                                        state="readonly", width=12, font=self._get_font_with_fallback(9),
                                        style='TCombobox')
        tooltip_font_combo.pack(side=tk.LEFT, padx=(0, Constants.WIDGET_PADDING))
        
        # Font size selection
        ttk.Label(tooltips_frame, text="Size:", font=self._get_font_with_fallback(9), 
                 style='TLabel').pack(side=tk.LEFT, padx=(0, 5))
        tooltip_size_entry = ttk.Entry(tooltips_frame, textvariable=self.tooltip_font_size_var, 
                                     width=Constants.ENTRY_WIDTH, font=self._get_font_with_fallback(9))
        tooltip_size_entry.pack(side=tk.LEFT)
    
    def _create_content_section(self, parent: ttk.Frame) -> None:
        """Create main content section with entries list and details."""
        content_frame = ttk.Frame(parent, style='TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, pady=15)
        
        self._create_entries_section(content_frame)
        ttk.Separator(content_frame, orient='vertical').pack(side=tk.LEFT, fill=tk.Y, padx=15)
        self._create_details_section(content_frame)
    
    def _create_entries_section(self, parent: ttk.Frame) -> None:
        """Create entries list section."""
        left_frame = ttk.Frame(parent, style='TFrame', padding=(0, 0, 15, 0))
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 15))
        
        ttk.Label(left_frame, text="Dock Entries", 
                 font=self._get_font_with_fallback(Constants.HEADER_FONT_SIZE, "bold"), 
                 style='TLabel').pack(anchor=tk.W, pady=(0, 15))
        
        # Entries listbox with scrollbar
        self._create_entries_listbox(left_frame)
        self._create_list_buttons(left_frame)
    
    def _create_entries_listbox(self, parent: ttk.Frame) -> None:
        """Create the entries listbox with scrollbar."""
        entries_frame = ttk.Frame(parent, style='TFrame')
        entries_frame.pack(fill=tk.BOTH, expand=True)
        
        self.entries_listbox = tk.Listbox(entries_frame, 
                                         bg=Constants.LISTBOX_BG, 
                                         fg=Constants.FG_COLOR, 
                                         selectbackground=Constants.SELECTED_BG, 
                                         font=self._get_font_with_fallback(Constants.LISTBOX_FONT_SIZE),
                                         borderwidth=1, 
                                         relief=tk.SOLID,
                                         activestyle='none',
                                         highlightthickness=0)
        self.entries_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.entries_listbox.bind('<<ListboxSelect>>', self.on_select_entry)
        
        scrollbar = ttk.Scrollbar(entries_frame, orient=tk.VERTICAL, command=self.entries_listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.entries_listbox.config(yscrollcommand=scrollbar.set)
    
    def _create_list_buttons(self, parent: ttk.Frame) -> None:
        """Create buttons for list management."""
        list_buttons_frame = ttk.Frame(parent, style='TFrame')
        list_buttons_frame.pack(fill=tk.X, pady=Constants.WIDGET_PADDING)
        
        # Top row buttons
        top_buttons = ttk.Frame(list_buttons_frame, style='TFrame')
        top_buttons.pack(anchor='center', pady=(0, 5))
        
        ttk.Button(top_buttons, text="Add Entry", command=self.add_entry, 
                  width=Constants.LISTBOX_SELECTION_WIDTH).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(top_buttons, text="Add Separator", command=self.add_separator, 
                  width=Constants.LISTBOX_SELECTION_WIDTH).pack(side=tk.LEFT, padx=(2, 0))
        
        # Middle row buttons
        middle_buttons = ttk.Frame(list_buttons_frame, style='TFrame')
        middle_buttons.pack(anchor='center', pady=(0, 5))
        
        ttk.Button(middle_buttons, text="Move Up", command=lambda: self.move_entry(-1), 
                  width=Constants.LISTBOX_SELECTION_WIDTH).pack(side=tk.LEFT, padx=(0, 2))
        ttk.Button(middle_buttons, text="Move Down", command=lambda: self.move_entry(1), 
                  width=Constants.LISTBOX_SELECTION_WIDTH).pack(side=tk.LEFT, padx=(2, 0))
        
        # Bottom row button
        bottom_buttons = ttk.Frame(list_buttons_frame, style='TFrame')
        bottom_buttons.pack(anchor='center')
        
        ttk.Button(bottom_buttons, text="Remove Entry", command=self.remove_entry, 
                  style='Danger.TButton', width=Constants.UPDATE_BUTTON_WIDTH).pack()
    
    def _create_details_section(self, parent: ttk.Frame) -> None:
        """Create entry details section."""
        right_frame = ttk.Frame(parent, style='TFrame')
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(15, 0))
        
        ttk.Label(right_frame, text="Entry Details", 
                 font=self._get_font_with_fallback(Constants.HEADER_FONT_SIZE, "bold"), 
                 style='TLabel').pack(anchor=tk.W, pady=(0, 15))
        
        details_frame = ttk.Frame(right_frame, style='TFrame')
        details_frame.pack(fill=tk.BOTH, expand=True)
        
        # Entry fields
        self._create_entry_field(details_frame, "Name:", "name_entry")
        self._create_entry_field_with_browse(details_frame, "App Path:", "app_path_entry", 
                                           self.browse_app_path, "Browse")
        self._create_entry_field_with_browse(details_frame, "Icon Path:", "icon_path_entry", 
                                           self.browse_icon_path, "Browse")
        
        # Update button
        self._create_update_button(details_frame)
    
    def _create_entry_field(self, parent: ttk.Frame, label_text: str, entry_attr: str) -> None:
        """Create a simple entry field with label."""
        frame = ttk.Frame(parent, style='TFrame')
        frame.pack(fill=tk.X, pady=Constants.WIDGET_PADDING)
        
        ttk.Label(frame, text=label_text, width=12, 
                 font=self._get_font_with_fallback(Constants.ENTRY_FONT_SIZE, "bold"), 
                 style='TLabel').pack(side=tk.LEFT)
        entry = ttk.Entry(frame, font=self._get_font_with_fallback(Constants.ENTRY_FONT_SIZE))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))
        setattr(self, entry_attr, entry)
    
    def _create_entry_field_with_browse(self, parent: ttk.Frame, label_text: str, 
                                       entry_attr: str, browse_command, button_text: str) -> None:
        """Create an entry field with browse button."""
        frame = ttk.Frame(parent, style='TFrame')
        frame.pack(fill=tk.X, pady=Constants.WIDGET_PADDING)
        
        ttk.Label(frame, text=label_text, width=12, 
                 font=self._get_font_with_fallback(Constants.ENTRY_FONT_SIZE, "bold"), 
                 style='TLabel').pack(side=tk.LEFT)
        entry = ttk.Entry(frame, font=self._get_font_with_fallback(Constants.ENTRY_FONT_SIZE))
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 5))
        ttk.Button(frame, text=button_text, command=browse_command, width=10).pack(side=tk.RIGHT)
        setattr(self, entry_attr, entry)
    
    def _create_update_button(self, parent: ttk.Frame) -> None:
        """Create the update entry button."""
        update_frame = ttk.Frame(parent, style='TFrame')
        update_frame.pack(fill=tk.X, pady=(20, Constants.WIDGET_PADDING))
        
        self.update_button = ttk.Button(update_frame, text="Update Entry", command=self.update_entry, 
                                       style='Success.TButton', padding=Constants.WIDGET_PADDING)
        self.update_button.pack(fill=tk.X)
    
    def _create_bottom_section(self, parent: ttk.Frame) -> None:
        """Create bottom section with save and refresh buttons."""
        bottom_frame = ttk.Frame(parent, style='TFrame')
        bottom_frame.pack(fill=tk.X)
        
        ttk.Button(bottom_frame, text="Save Changes", command=self.save_ini, 
                  style='Accent.TButton', padding=Constants.WIDGET_PADDING).pack(
                      side=tk.LEFT, fill=tk.X, expand=True, padx=(0, Constants.WIDGET_PADDING))
        ttk.Button(bottom_frame, text="Refresh Dock", command=self.refresh_rainmeter, 
                  padding=Constants.WIDGET_PADDING).pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    def _create_status_bar(self) -> None:
        """Create status bar."""
        status_frame = ttk.Frame(self.master, style='TFrame')
        status_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        status_bar = ttk.Label(status_frame, textvariable=self.status_message, 
                              background=Constants.STATUS_BG, foreground="#aaaaaa",
                              relief=tk.FLAT, anchor=tk.W, padding=(Constants.WIDGET_PADDING, 5))
        status_bar.pack(fill=tk.X)
    
    def setup_keyboard_shortcuts(self) -> None:
        """Setup keyboard shortcuts."""
        shortcuts = {
            "<Control-s>": lambda event: self.save_ini(),
            "<Control-n>": lambda event: self.add_entry(),
            "<Control-d>": lambda event: self.remove_entry(),
            "<Control-Up>": lambda event: self.move_entry(-1),
            "<Control-Down>": lambda event: self.move_entry(1),
            "<Control-Shift-n>": lambda event: self.create_new_dock()
        }
        
        for key, command in shortcuts.items():
            self.master.bind(key, command)
    
    def set_status(self, message: str) -> None:
        """Set status message."""
        self.status_message.set(message)
        self.master.update_idletasks()
    
    # Event handlers
    def on_radius_scale_change(self, value: str) -> None:
        """Handle radius slider change."""
        radius = int(float(value))
        self.corner_radius_var.set(radius)
        self.set_status(f"Corner radius changed to {radius} - Click 'Save Changes' to apply")
    
    def on_radius_entry_change(self, event=None) -> None:
        """Handle radius entry change."""
        try:
            value = int(self.corner_radius_var.get())
            value = max(Constants.MIN_CORNER_RADIUS, min(Constants.MAX_CORNER_RADIUS, value))
            self.corner_radius_var.set(value)
        except ValueError:
            self.corner_radius_var.set(Constants.DEFAULT_CORNER_RADIUS)
    
    def on_size_scale_change(self, value: str) -> None:
        """Handle size slider change."""
        size = int(float(value))
        self.dock_size_var.set(size)
        self.set_status(f"Icon size changed to {size} - Click 'Save Changes' to apply")
    
    def on_size_entry_change(self, event=None) -> None:
        """Handle size entry change."""
        try:
            value = int(self.dock_size_var.get())
            value = max(Constants.MIN_DOCK_SIZE, min(Constants.MAX_DOCK_SIZE, value))
            self.dock_size_var.set(value)
        except ValueError:
            self.dock_size_var.set(Constants.DEFAULT_DOCK_SIZE)
    
    def on_orientation_change(self, event=None) -> None:
        """Handle orientation change."""
        orientation = self.orientation_var.get()
        self.set_status(f"Dock orientation changed to {orientation} - Click 'Save Changes' to apply")
    
    def choose_color(self) -> None:
        """Open color picker dialog."""
        current_hex = f"#{self.current_color[0]:02x}{self.current_color[1]:02x}{self.current_color[2]:02x}"
        color = colorchooser.askcolor(color=current_hex, title="Choose Dock Background Color")
        if color[0]:
            self.current_color = tuple(int(c) for c in color[0])
            self._update_color_display()
            self.set_status(f"Color selected: RGB{self.current_color}")
    
    def _update_color_display(self) -> None:
        """Update color preview and RGB entry."""
        hex_color = f"#{self.current_color[0]:02x}{self.current_color[1]:02x}{self.current_color[2]:02x}"
        self.color_preview.configure(bg=hex_color)
        self.rgb_entry.delete(0, tk.END)
        self.rgb_entry.insert(0, f"{self.current_color[0]},{self.current_color[1]},{self.current_color[2]}")
    
    def on_manual_color_change(self, event=None) -> None:
        """Handle manual RGB entry change."""
        try:
            rgb_text = self.rgb_entry.get().strip()
            parts = [int(x.strip()) for x in rgb_text.split(',')]
            if len(parts) == 3 and all(0 <= val <= 255 for val in parts):
                self.current_color = tuple(parts)
                hex_color = f"#{parts[0]:02x}{parts[1]:02x}{parts[2]:02x}"
                self.color_preview.configure(bg=hex_color)
                self.set_status(f"Manual color set: RGB{self.current_color}")
            else:
                raise ValueError("Invalid RGB values")
        except (ValueError, AttributeError):
            self._update_color_display()
            self.set_status("Invalid RGB format. Use: R,G,B (0-255)")
    
    def on_dock_selected(self, event) -> None:
        """Handle dock selection change."""
        selected_dock = self.current_dock.get()
        self.load_dock(selected_dock)
    
    def on_select_entry(self, event) -> None:
        """Handle entry selection in listbox."""
        selection = self.entries_listbox.curselection()
        if selection:
            index = selection[0]
            if self.selected_entry_index is not None:
                self.entries_listbox.itemconfig(self.selected_entry_index, bg=Constants.LISTBOX_BG)
            
            self.selected_entry_index = index
            self.entries_listbox.itemconfig(index, bg=Constants.SELECTED_BG)
            
            entry = self.entries[index]
            self._populate_entry_fields(entry)
            self.set_status(f"Selected entry: {entry.name}")
    
    def _populate_entry_fields(self, entry: DockEntry) -> None:
        """Populate entry detail fields with data from selected entry."""
        self.name_entry.delete(0, tk.END)
        self.name_entry.insert(0, entry.name)
        
        self.app_path_entry.delete(0, tk.END)
        self.app_path_entry.insert(0, entry.app_path)
        
        self.icon_path_entry.delete(0, tk.END)
        self.icon_path_entry.insert(0, entry.icon_path)
    
    # File dialog methods
    def browse_app_path(self) -> None:
        """Browse for application executable."""
        file_path = filedialog.askopenfilename(
            title="Select Application",
            filetypes=Constants.EXECUTABLE_FILTER
        )
        if file_path:
            self.app_path_entry.delete(0, tk.END)
            self.app_path_entry.insert(0, f'"{file_path}"')
            self.set_status(f"Selected application: {file_path}")
    
    def browse_icon_path(self) -> None:
        """Browse for icon file."""
        file_path = filedialog.askopenfilename(
            title="Select Icon",
            filetypes=Constants.IMAGE_FILTER
        )
        if file_path:
            self.icon_path_entry.delete(0, tk.END)
            self.icon_path_entry.insert(0, file_path)
            self.set_status(f"Selected icon: {file_path}")
    
    # Dock management methods
    def load_dock(self, dock_name: str) -> None:
        """Load a specific dock configuration."""
        self.set_status(f"Loading {dock_name}...")
        
        if not DockValidator.validate_dock_name(dock_name):
            messagebox.showerror("Error", f"Invalid dock name format: {dock_name}")
            self.set_status(f"Error: Invalid dock name format: {dock_name}")
            return
        
        dock_number = dock_name[len(Constants.CONFIG_PREFIX):]
        self.ini_file = os.path.join(self.base_dir, f"{Constants.DOCK_PREFIX}{dock_number}", 
                                   f"{dock_name}.ini")
        
        try:
            self._load_entries_from_ini()
            self._load_appearance_settings()
            self.set_status(f"Loaded {len(self.entries)} entries from {dock_name}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load {dock_name}: {str(e)}")
            self.set_status(f"Error loading {dock_name}: {str(e)}")
    
    def _load_entries_from_ini(self) -> None:
        """Load entries from the current INI file."""
        dock_entries = self.config_manager.load_entries_from_ini(self.ini_file)
        
        self.entries = dock_entries
        self.entries_listbox.delete(0, tk.END)
        
        for entry in self.entries:
            self.entries_listbox.insert(tk.END, entry.name)
            if entry.is_separator:
                self.entries_listbox.itemconfig(tk.END, fg=Constants.SEPARATOR_COLOR)
        
        self._clear_entry_fields()
    
    def _load_appearance_settings(self) -> None:
        """Load appearance settings from the current dock."""
        settings = self.config_manager.load_appearance_settings(self.ini_file)
        
        self.corner_radius_var.set(settings['corner_radius'])
        self.dock_size_var.set(settings['dock_size'])
        self.current_color = settings['current_color']
        self._update_color_display()
        self.opacity_var.set(settings['opacity'])
        self.orientation_var.set(settings['orientation'])
        self.double_click_var.set(settings['double_click'])
        self.tooltips_var.set(settings['tooltips'])
        self.tooltip_delay_var.set(settings['tooltip_delay'])
        self.tooltip_font_var.set(settings['tooltip_font'])
        self.tooltip_font_size_var.set(settings['tooltip_font_size'])
    
    def _clear_entry_fields(self) -> None:
        """Clear entry detail fields."""
        self.selected_entry_index = None
        for entry_field in [self.name_entry, self.app_path_entry, self.icon_path_entry]:
            entry_field.delete(0, tk.END)
    
    def refresh_dock_list(self) -> None:
        """Refresh available docks list."""
        current_selection = self.current_dock.get()
        self.available_docks = self.config_manager.find_available_docks()
        self.dock_combo['values'] = self.available_docks
        
        if current_selection in self.available_docks:
            self.current_dock.set(current_selection)
        elif self.available_docks:
            self.current_dock.set(self.available_docks[0])
        else:
            self.current_dock.set("")
    
    def create_new_dock(self) -> None:
        """Create a new dock from template."""
        try:
            dock_number = self.config_manager.get_next_dock_number(self.available_docks)
            dock_file_name = self.config_manager.create_new_dock(dock_number)
            
            self.refresh_dock_list()
            self.current_dock.set(dock_file_name)
            self.load_dock(dock_file_name)
            
            messagebox.showinfo("Success", f"Successfully created 'Lucky Dock {dock_number}' ({Constants.DOCK_PREFIX}{dock_number})")
            self.set_status(f"Created new dock: Lucky Dock {dock_number}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create dock: {str(e)}")
            self.set_status(f"Error creating dock: {str(e)}")
    
    def delete_dock(self) -> None:
        """Delete the currently selected dock."""
        if not self.current_dock.get():
            messagebox.showerror("Error", Constants.ERROR_NO_DOCK)
            self.set_status(f"Error: {Constants.ERROR_NO_DOCK}")
            return
        
        current_dock_name = self.current_dock.get()
        
        if not DockValidator.validate_dock_name(current_dock_name):
            messagebox.showerror("Error", Constants.ERROR_INVALID_DOCK)
            self.set_status(f"Error: {Constants.ERROR_INVALID_DOCK}")
            return
        
        dock_number = current_dock_name[len(Constants.CONFIG_PREFIX):]
        dock_display_name = f"Lucky Dock {dock_number}"
        
        if not messagebox.askyesno("Confirm Delete", 
                                  f"Are you sure you want to delete '{dock_display_name}' ({Constants.DOCK_PREFIX}{dock_number})?\n\n"
                                  f"This will permanently delete all dock files and cannot be undone."):
            self.set_status("Dock deletion cancelled")
            return
        
        try:
            self.set_status("Unloading dock from Rainmeter...")
            self.rainmeter.unload_skin(dock_number)
            
            self.config_manager.delete_dock(dock_number)
            
            self.refresh_dock_list()
            self._handle_post_deletion()
            
            messagebox.showinfo("Success", f"Successfully deleted dock '{dock_display_name}'")
            self.set_status(f"Successfully deleted dock: {dock_display_name}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete dock: {str(e)}")
            self.set_status(f"Error deleting dock: {str(e)}")
    
    def _handle_post_deletion(self) -> None:
        """Handle UI state after dock deletion."""
        if self.available_docks:
            self.current_dock.set(self.available_docks[0])
            self.load_dock(self.available_docks[0])
        else:
            self.current_dock.set("")
            self.ini_file = None
            self.entries = []
            self.entries_listbox.delete(0, tk.END)
            self._clear_entry_fields()
    
    # Entry management methods
    def add_entry(self) -> None:
        """Add a new entry."""
        if not self.ini_file:
            messagebox.showerror("Error", Constants.ERROR_NO_DOCK)
            self.set_status(f"Error: {Constants.ERROR_NO_DOCK}")
            return
        
        new_entry = DockEntry("New Entry", '""', '""', False)
        self.entries.append(new_entry)
        self.entries_listbox.insert(tk.END, new_entry.name)
        self._select_last_entry()
        
        self._auto_save_and_refresh(Constants.STATUS_ADDED)
    
    def add_separator(self) -> None:
        """Add a new separator."""
        if not self.ini_file:
            messagebox.showerror("Error", Constants.ERROR_NO_DOCK)
            self.set_status(f"Error: {Constants.ERROR_NO_DOCK}")
            return
        
        separator_count = sum(1 for entry in self.entries if entry.is_separator)
        new_name = f"Separator_{separator_count + 1}"
        separator_icon = os.path.join(self.base_dir, Constants.RESOURCES_DIR, 
                                    Constants.IMAGES_DIR, Constants.SEPARATOR_ICON)
        
        new_entry = DockEntry(new_name, '""', separator_icon, True)
        self.entries.append(new_entry)
        self.entries_listbox.insert(tk.END, new_name)
        self.entries_listbox.itemconfig(tk.END, fg=Constants.SEPARATOR_COLOR)
        self._select_last_entry()
        
        self._auto_save_and_refresh("Added new separator")
    
    def _select_last_entry(self) -> None:
        """Select the last entry in the listbox."""
        self.entries_listbox.selection_clear(0, tk.END)
        self.entries_listbox.selection_set(tk.END)
        self.entries_listbox.see(tk.END)
        self.on_select_entry(None)
    
    def update_entry(self) -> None:
        """Update the selected entry."""
        if self.selected_entry_index is None:
            messagebox.showerror("Error", Constants.ERROR_NO_ENTRY)
            self.set_status(f"Error: {Constants.ERROR_NO_ENTRY}")
            return
        
        name = self.name_entry.get().strip()
        if not name:
            messagebox.showerror("Validation Error", Constants.ERROR_EMPTY_NAME)
            self.set_status(f"Error: {Constants.ERROR_EMPTY_NAME}")
            return
        
        is_separator = name.startswith("Separator_")
        
        updated_entry = DockEntry(
            name=name,
            app_path=self.app_path_entry.get(),
            icon_path=self.icon_path_entry.get(),
            is_separator=is_separator
        )
        
        self.entries[self.selected_entry_index] = updated_entry
        
        self.entries_listbox.delete(self.selected_entry_index)
        self.entries_listbox.insert(self.selected_entry_index, name)
        
        color = Constants.SEPARATOR_COLOR if is_separator else Constants.FG_COLOR
        self.entries_listbox.itemconfig(self.selected_entry_index, fg=color)
        self.entries_listbox.selection_set(self.selected_entry_index)
        
        self._auto_save_and_refresh(f"Updated entry: {name}")
    
    def remove_entry(self) -> None:
        """Remove the selected entry."""
        if self.selected_entry_index is None:
            messagebox.showerror("Error", Constants.ERROR_NO_ENTRY)
            self.set_status(f"Error: {Constants.ERROR_NO_ENTRY}")
            return
        
        entry_name = self.entries[self.selected_entry_index].name
        if not messagebox.askyesno("Confirm Delete", f"Are you sure you want to delete '{entry_name}'?"):
            self.set_status("Delete cancelled")
            return
        
        index = self.selected_entry_index
        del self.entries[index]
        self.entries_listbox.delete(index)
        self.selected_entry_index = None
        
        self._handle_post_removal(index)
        self._auto_save_and_refresh(f"Removed entry '{entry_name}'")
    
    def _handle_post_removal(self, removed_index: int) -> None:
        """Handle UI state after entry removal."""
        if self.entries:
            new_index = min(removed_index, len(self.entries) - 1)
            self.entries_listbox.selection_set(new_index)
            self.on_select_entry(None)
        else:
            self._clear_entry_fields()
    
    def move_entry(self, direction: int) -> None:
        """Move the selected entry up or down."""
        if self.selected_entry_index is None:
            messagebox.showerror("Error", Constants.ERROR_NO_ENTRY)
            self.set_status(f"Error: {Constants.ERROR_NO_ENTRY}")
            return
        
        index = self.selected_entry_index
        new_index = index + direction
        
        if 0 <= new_index < len(self.entries):
            moved_entry_name = self.entries[index].name
            
            # Swap entries
            self.entries[index], self.entries[new_index] = self.entries[new_index], self.entries[index]
            
            # Update listbox
            self.entries_listbox.delete(index)
            self.entries_listbox.insert(new_index, self.entries[new_index].name)
            
            # Update colors
            color = Constants.SEPARATOR_COLOR if self.entries[new_index].is_separator else Constants.FG_COLOR
            self.entries_listbox.itemconfig(new_index, fg=color)
            
            self.entries_listbox.selection_clear(0, tk.END)
            self.entries_listbox.selection_set(new_index)
            self.entries_listbox.see(new_index)
            self.selected_entry_index = new_index
            
            direction_text = "up" if direction < 0 else "down"
            self._auto_save_and_refresh(f"Moved '{moved_entry_name}' {direction_text}")
        else:
            direction_text = "up" if direction < 0 else "down"
            self.set_status(f"Cannot move entry further {direction_text}")
    
    def _auto_save_and_refresh(self, action_description: str) -> None:
        """Auto-save and refresh after entry changes."""
        self.set_status(f"{action_description} - Saving and refreshing...")
        try:
            self._save_ini_silent()
            self._refresh_rainmeter_silent()
            self.set_status(f"{action_description} and applied successfully")
        except Exception as e:
            self.set_status(f"{action_description} but failed to apply changes: {str(e)}")
    
    # Save and validation methods
    def save_ini(self) -> None:
        """Save INI file with user feedback."""
        if not self.ini_file:
            messagebox.showerror("Error", Constants.ERROR_NO_DOCK)
            self.set_status(f"Error: {Constants.ERROR_NO_DOCK}")
            return
        
        self.set_status(Constants.STATUS_SAVING)
        
        if not DockValidator.validate_entries(self.entries, show_dialogs=True):
            self.set_status("Save cancelled due to validation errors")
            return
        
        try:
            self._save_entries_to_ini()
            self.refresh_rainmeter()
            messagebox.showinfo("Success", "Changes saved and dock refreshed successfully!")
            self.set_status("Changes saved successfully")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save INI file: {str(e)}")
            self.set_status(f"Error saving INI file: {str(e)}")
    
    def _save_ini_silent(self) -> None:
        """Save INI file silently without user dialogs."""
        if not self.ini_file:
            raise Exception(Constants.ERROR_NO_DOCK)
        
        if not DockValidator.validate_entries(self.entries, show_dialogs=False):
            raise Exception("Validation failed")
        
        self._save_entries_to_ini()
    
    def _save_entries_to_ini(self) -> None:
        """Save entries and settings to the INI file."""
        appearance_settings = self._get_current_appearance_settings()
        self.config_manager.save_entries_to_ini(self.ini_file, self.entries, appearance_settings)
    
    def _get_current_appearance_settings(self) -> Dict[str, Any]:
        """Get current appearance settings from GUI variables."""
        return {
            'corner_radius': self.corner_radius_var.get(),
            'dock_size': self.dock_size_var.get(),
            'current_color': self.current_color,
            'opacity': self.opacity_var.get(),
            'orientation': self.orientation_var.get(),
            'double_click': self.double_click_var.get(),
            'tooltips': self.tooltips_var.get(),
            'tooltip_delay': self.tooltip_delay_var.get(),
            'tooltip_font': self.tooltip_font_var.get(),
            'tooltip_font_size': self.tooltip_font_size_var.get()
        }
    
    def refresh_rainmeter(self) -> None:
        """Refresh Rainmeter skin with user feedback."""
        if not self.ini_file:
            messagebox.showerror("Error", Constants.ERROR_NO_DOCK)
            self.set_status(f"Error: {Constants.ERROR_NO_DOCK}")
            return
        
        try:
            self._refresh_rainmeter_silent()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to refresh Rainmeter skin: {str(e)}")
            self.set_status(f"Error refreshing Rainmeter skin: {str(e)}")
    
    def _refresh_rainmeter_silent(self) -> None:
        """Refresh Rainmeter skin silently."""
        if not self.ini_file:
            raise Exception(Constants.ERROR_NO_DOCK)
        
        dock_name = os.path.basename(self.ini_file).split('.')[0]
        
        if not DockValidator.validate_dock_name(dock_name):
            raise Exception(f"Invalid dock name format: {dock_name}")
        
        dock_number = dock_name[len(Constants.CONFIG_PREFIX):]
        
        if not self.rainmeter.refresh_skin(dock_number):
            raise Exception("Failed to refresh Rainmeter skin")
        
        self.set_status(f"Activated and refreshed: {Constants.CONFIG_PREFIX}\\{Constants.DOCK_PREFIX}{dock_number}")


def load_application_icon(root: tk.Tk) -> None:
    """Load application icon if available."""
    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(script_dir, Constants.ICON_FILENAME)
        if os.path.exists(icon_path):
            root.iconbitmap(icon_path)
    except Exception:
        pass  # Icon loading is optional


def parse_command_line_arguments() -> Optional[str]:
    """Parse command line arguments to get dock number."""
    dock_number = None
    if len(sys.argv) > 1:
        try:
            dock_arg = sys.argv[1]
            if dock_arg.isdigit():
                dock_number = dock_arg
            elif dock_arg.startswith(Constants.CONFIG_PREFIX) and dock_arg[len(Constants.CONFIG_PREFIX):].isdigit():
                dock_number = dock_arg[len(Constants.CONFIG_PREFIX):]
        except Exception:
            pass  # Invalid arguments are ignored
    return dock_number


def main() -> None:
    """Main application entry point."""
    # Parse command line arguments
    dock_number = parse_command_line_arguments()
    
    # Create and configure main window
    root = tk.Tk()
    load_application_icon(root)
    
    # Create and run application
    app = LuckyDockManager(root, dock_number)
    root.mainloop()


if __name__ == "__main__":
    main()