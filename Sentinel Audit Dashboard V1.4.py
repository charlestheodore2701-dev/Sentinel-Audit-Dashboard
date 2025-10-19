# -*- coding: utf-8 -*-
"""
Sentinel Audit Dashboard V1.4 - Enhanced Version with Equipment Notes
Created on Mon Oct 13 13:21:52 2025
@author: charl

New Features in V1.4:
- Equipment notes system for tracking maintenance and changes
- Equipment search/filter functionality with highlighting
- Database statistics view
- Enhanced error logging
- Improved UI responsiveness
- Better error handling and user feedback
- Visual highlighting for equipment with notes
"""

# Standard imports
import tkinter as tk
from tkinter import filedialog, ttk, messagebox, simpledialog
from tkcalendar import DateEntry
import pandas as pd
import sqlite3
import json
import os
import threading
from PIL import Image, ImageTk 
from datetime import date, timedelta, datetime
import sys 
import contextlib
import shutil
import traceback

# Matplotlib imports for charting
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.dates as mdates
import matplotlib.ticker as ticker
plt.style.use('ggplot')

# --- CONSTANTS ---
EXCEL_DATE_ORIGIN = '1899-12-30'
DATETIME_FORMAT = '%Y-%m-%d %H:%M:%S'
DEFAULT_DATE_RANGE_DAYS = 30 
LOG_FILENAME = "sentinel_audit_log.txt"

# -----------------------------------------------
# UTILITY FUNCTIONS
# -----------------------------------------------

def run_in_thread(target_func):
    """Starts a target function in a new thread."""
    threading.Thread(target=target_func, daemon=True).start()
    
def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
    
def get_script_path():
    """Gets the directory where the main script file resides."""
    return os.path.dirname(os.path.abspath(__file__))

def log_message(message, level="INFO"):
    """Log messages to file with timestamp"""
    try:
        log_path = os.path.join(get_script_path(), LOG_FILENAME)
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with open(log_path, "a", encoding="utf-8") as log_file:
            log_file.write(f"[{timestamp}] [{level}] {message}\n")
    except Exception as e:
        print(f"Logging error: {e}")

# --- PATH LOGIC ---
APP_DATA_DIR = get_script_path()
SHAFT_JSON_FILENAME = "shaft_list.json"
SHAFT_JSON_PATH = os.path.join(APP_DATA_DIR, SHAFT_JSON_FILENAME)


def load_shaft_databases():
    """Loads site configuration with robust error handling"""
    if os.path.exists(SHAFT_JSON_PATH):
        try:
            with open(SHAFT_JSON_PATH, "r", encoding="utf-8") as f:
                return json.load(f)
        except (json.JSONDecodeError, UnicodeDecodeError) as err:
            log_message(f"UTF-8 read failed: {err}. Trying fallback encoding...", "WARNING")
            
            try:
                with open(SHAFT_JSON_PATH, "r", encoding="latin-1") as f:
                    content = f.read()
                    shaft_data = json.loads(content) 
                    log_message("Successfully read config with latin-1 fallback", "INFO")
                    save_shaft_databases(shaft_data) 
                    return shaft_data
            except Exception as fallback_err:
                log_message(f"Fallback to latin-1 failed: {fallback_err}", "ERROR")
                
    # Default configuration
    default_config = {
        "K3": "sentinel_k3.db",
        "Saffy": "sentinel_saffy.db",
        "Rowland": "sentinel_rowland.db",
        "EPL3": "sentinel_epl3.db",
        "Glencore": "sentinel_glencore.db"
    }
    save_shaft_databases(default_config)
    log_message("Created default configuration", "INFO")
    return default_config


def save_shaft_databases(shaft_dict):
    """Saves the shaft configuration to JSON file with UTF-8 encoding and backup"""
    try:
        # Create backup of existing config
        if os.path.exists(SHAFT_JSON_PATH):
            backup_path = SHAFT_JSON_PATH + '.backup'
            shutil.copy2(SHAFT_JSON_PATH, backup_path)
        
        # Save new config
        with open(SHAFT_JSON_PATH, "w", encoding="utf-8") as f:
            json.dump(shaft_dict, f, indent=4, ensure_ascii=False)
        
        log_message("Configuration saved successfully", "INFO")
            
    except Exception as err:
        log_message(f"Error saving shaft list: {err}", "ERROR")
        # Try to restore from backup
        backup_path = SHAFT_JSON_PATH + '.backup'
        if os.path.exists(backup_path):
            shutil.copy2(backup_path, SHAFT_JSON_PATH)
            log_message("Restored configuration from backup", "WARNING")


# -----------------------------------------------
# MAIN APPLICATION CLASS
# -----------------------------------------------
class SentinelDashboard(tk.Tk):
    def __init__(self):
        super().__init__()
        
        # Setup State Variables
        self.shaft_databases_cache = load_shaft_databases()
        self.selected_shaft = tk.StringVar()
        if self.shaft_databases_cache:
            self.selected_shaft.set(list(self.shaft_databases_cache.keys())[0])
        else:
            self.selected_shaft.set("") 
        
        # Initialize Widget Attributes
        self.shaft_dropdown = None
        self.from_date = None
        self.to_date = None
        self.dashboard_tree = None
        self.logo_photo = None 
        self.status_label = None 
        self.progress_var = tk.DoubleVar()
        self.progress_bar = None
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_change)
        
        # Dashboard caching
        self._dashboard_cache = None
        self._cache_key = None
        self._full_data = None

        # Setup UI
        self.title("Sentinel Audit Dashboard V1.4")
        self.geometry("1200x800")
        self.configure(bg='#f0f0f0')
        
        # Define color scheme
        self.colors = {
            'bg': '#f0f0f0',
            'primary': '#2c3e50',
            'secondary': '#3498db',
            'success': '#27ae60',
            'danger': '#e74c3c',
            'warning': '#f39c12',
            'text': '#2c3e50',
            'info': '#3498db'
        }
        
        self.setup_ui() 
        
        # Only run init_db and refresh if a shaft is selected
        if self.selected_shaft.get():
            self.init_db()
            self.refresh_dashboard_table()
        
        log_message("Application started", "INFO")

    def setup_ui(self):
        """Setup the user interface"""
        # --- Header Frame with Logo and Status ---
        header_frame = tk.Frame(self, bg='white', relief=tk.FLAT, bd=0)
        header_frame.pack(fill='x', padx=0, pady=0)
        
        header_inner = tk.Frame(header_frame, bg='white')
        header_inner.pack(fill='x', padx=20, pady=15)
        
        # Logo Setup
        try:
            logo_path = resource_path("Schauenburg logo.png")
            if os.path.exists(logo_path):
                logo_img = Image.open(logo_path)
                logo_img = logo_img.resize((350, 100), Image.Resampling.LANCZOS)
                self.logo_photo = ImageTk.PhotoImage(logo_img) 
                logo_label = tk.Label(header_inner, image=self.logo_photo, bg='white')
                logo_label.image = self.logo_photo 
                logo_label.pack(side="left", padx=(0, 15))
        except Exception as err:
            log_message(f"Error loading logo: {err}", "WARNING")
        
        # Title Label
        title_label = tk.Label(header_inner, text="Sentinel Audit Dashboard V1.4", 
                              font=('Segoe UI', 18, 'bold'), 
                              bg='white', fg=self.colors['primary'])
        title_label.pack(side="left", padx=20)
            
        # Status Label
        self.status_label = tk.Label(header_inner, text="Ready", 
                                     font=('Segoe UI', 10), 
                                     fg=self.colors['success'], bg='white')
        self.status_label.pack(side="right", padx=10)
        
        # Separator
        separator = tk.Frame(self, height=2, bg='#ddd')
        separator.pack(fill='x')
        
        # Main content frame
        content_frame = tk.Frame(self, bg=self.colors['bg'])
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # --- Control Buttons Frame ---
        control_frame = tk.Frame(content_frame, bg=self.colors['bg'])
        control_frame.pack(fill='x', pady=(0, 15))
        
        btn_style = {'font': ('Segoe UI', 9), 'relief': tk.FLAT, 'bd': 0, 
                    'padx': 15, 'pady': 8, 'cursor': 'hand2'}
        
        self.btn_add_site = tk.Button(control_frame, text="➕ Add Site", 
                                      bg=self.colors['success'], fg='white',
                                      activebackground='#229954',
                                      command=self.add_new_site, **btn_style)
        self.btn_add_site.pack(side="left", padx=(0, 10))

        self.btn_remove_site = tk.Button(control_frame, text="➖ Remove Site", 
                                        bg=self.colors['danger'], fg='white',
                                        activebackground='#c0392b',
                                        command=self.remove_site, **btn_style)
        self.btn_remove_site.pack(side="left", padx=(0, 10))
        
        self.btn_reset_db = tk.Button(control_frame, text="🔄 Reset Database", 
                                      bg=self.colors['warning'], fg='white',
                                      activebackground='#d68910',
                                      command=self.reset_database, **btn_style)
        self.btn_reset_db.pack(side="left", padx=(0, 10))
        
        self.btn_stats = tk.Button(control_frame, text="📊 DB Stats", 
                                   bg=self.colors['info'], fg='white',
                                   activebackground='#2980b9',
                                   command=self.show_database_stats, **btn_style)
        self.btn_stats.pack(side="left")
        
        # --- Shaft Selection Frame ---
        shaft_frame = tk.Frame(content_frame, bg=self.colors['bg'])
        shaft_frame.pack(fill='x', pady=(0, 15))
        
        shaft_label = tk.Label(shaft_frame, text="Select Site:", 
                              font=('Segoe UI', 10, 'bold'),
                              bg=self.colors['bg'], fg=self.colors['text'])
        shaft_label.pack(side='left', padx=(0, 10))
        
        # Style for combobox
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('Custom.TCombobox', 
                       fieldbackground='white',
                       background=self.colors['secondary'],
                       foreground=self.colors['text'],
                       arrowcolor=self.colors['secondary'])
        
        self.shaft_dropdown = ttk.Combobox(shaft_frame, textvariable=self.selected_shaft, 
                                           values=list(self.shaft_databases_cache.keys()), 
                                           state="readonly", width=25,
                                           font=('Segoe UI', 10),
                                           style='Custom.TCombobox')
        self.shaft_dropdown.pack(side='left', padx=(0, 20))
        self.shaft_dropdown.bind("<<ComboboxSelected>>", 
                                lambda _: [self.init_db(), self.refresh_dashboard_table()])

        # Search Box
        search_label = tk.Label(shaft_frame, text="🔍 Search Equipment:", 
                               font=('Segoe UI', 10, 'bold'),
                               bg=self.colors['bg'], fg=self.colors['text'])
        search_label.pack(side='left', padx=(0, 10))
        
        self.search_entry = tk.Entry(shaft_frame, textvariable=self.search_var,
                                     font=('Segoe UI', 10), width=20)
        self.search_entry.pack(side='left', padx=(0, 10))
        
        # Clear search button
        self.btn_clear_search = tk.Button(shaft_frame, text="✖", 
                                         bg=self.colors['danger'], fg='white',
                                         activebackground='#c0392b',
                                         command=self.clear_search,
                                         font=('Segoe UI', 8), relief=tk.FLAT, 
                                         bd=0, padx=8, pady=4, cursor='hand2')
        self.btn_clear_search.pack(side='left')

        # --- Action Buttons Frame ---
        action_frame = tk.Frame(content_frame, bg=self.colors['bg'])
        action_frame.pack(fill='x', pady=(0, 15))
        
        action_btn_style = {'font': ('Segoe UI', 9), 'relief': tk.FLAT, 'bd': 0,
                           'padx': 15, 'pady': 8, 'cursor': 'hand2',
                           'bg': self.colors['secondary'], 'fg': 'white',
                           'activebackground': '#2980b9'}
        
        self.btn_import = tk.Button(action_frame, text="📥 Import Test Report", 
                                    command=lambda: run_in_thread(self._threaded_import_excel),
                                    **action_btn_style)
        self.btn_import.pack(side="left", padx=(0, 10))

        self.btn_export = tk.Button(action_frame, text="📤 Export Dashboard", 
                                    command=lambda: run_in_thread(self._threaded_export_dashboard),
                                    **action_btn_style)
        self.btn_export.pack(side="left", padx=(0, 10))

        self.btn_fail_count_report = tk.Button(action_frame, text="📊 Fail Count Report", 
                                               command=lambda: run_in_thread(self._threaded_export_daily_fail_count_report),
                                               **action_btn_style)
        self.btn_fail_count_report.pack(side="left", padx=(0, 10))

        self.btn_interval_view = tk.Button(action_frame, text="⏱️ Test Counts", 
                                          command=self.show_test_counts_by_interval,
                                          **action_btn_style)
        self.btn_interval_view.pack(side="left", padx=(0, 10))
        
        self.btn_failure_trend = tk.Button(action_frame, text="📈 Failure Trend", 
                                          command=lambda: run_in_thread(self._threaded_show_failure_trend),
                                          **action_btn_style)
        self.btn_failure_trend.pack(side="left", padx=(0, 10))
        
        self.btn_consolidated_trend = tk.Button(action_frame, text="🌎 Consolidated Trend", 
                                                command=lambda: run_in_thread(self._threaded_show_consolidated_failure_trend),
                                                **action_btn_style)
        self.btn_consolidated_trend.pack(side="left", padx=(0, 10))

        self.btn_common_fail = tk.Button(action_frame, text="🔍 Common Failures", 
                                        command=self.show_most_common_failure,
                                        **action_btn_style)
        self.btn_common_fail.pack(side="left", padx=(0, 10))
        
        self.btn_notes = tk.Button(action_frame, text="📝 Equipment Notes", 
                                  command=self.view_selected_equipment_notes,
                                  **action_btn_style)
        self.btn_notes.pack(side="left")
        
        # --- Date Filter Frame ---
        date_frame = tk.Frame(content_frame, bg='white', relief=tk.FLAT, bd=0)
        date_frame.pack(fill='x', pady=(0, 15))
        
        date_inner = tk.Frame(date_frame, bg='white')
        date_inner.pack(fill='x', padx=15, pady=15)
        
        today = date.today()
        one_month_ago = today - timedelta(days=DEFAULT_DATE_RANGE_DAYS)
        
        # From Date
        from_label = tk.Label(date_inner, text="From Date:", 
                             font=('Segoe UI', 10, 'bold'),
                             bg='white', fg=self.colors['text'])
        from_label.pack(side="left", padx=(0, 10))
        
        self.from_date = DateEntry(date_inner, width=14, background=self.colors['secondary'], 
                                   foreground='white', borderwidth=0,
                                   font=('Segoe UI', 10))
        self.from_date.set_date(one_month_ago)
        self.from_date.pack(side="left", padx=(0, 20))

        # To Date
        to_label = tk.Label(date_inner, text="To Date:", 
                           font=('Segoe UI', 10, 'bold'),
                           bg='white', fg=self.colors['text'])
        to_label.pack(side="left", padx=(0, 10))
        
        self.to_date = DateEntry(date_inner, width=14, background=self.colors['secondary'], 
                                foreground='white', borderwidth=0,
                                font=('Segoe UI', 10))
        self.to_date.set_date(today)
        self.to_date.pack(side="left", padx=(0, 20))
        
        # Filter Button
        self.btn_filter = tk.Button(date_inner, text="Apply Filter", 
                                    command=lambda: run_in_thread(self._threaded_refresh_dashboard_table),
                                    font=('Segoe UI', 9, 'bold'), relief=tk.FLAT, bd=0,
                                    padx=20, pady=8, cursor='hand2',
                                    bg=self.colors['primary'], fg='white',
                                    activebackground='#1a252f')
        self.btn_filter.pack(side="left")

        # --- Progress Bar ---
        self.progress_bar = ttk.Progressbar(
            content_frame, 
            variable=self.progress_var,
            mode='determinate',
            length=400
        )

        # --- Legend Frame ---
        legend_frame = tk.Frame(content_frame, bg='white', relief=tk.FLAT, bd=0)
        legend_frame.pack(fill='x', pady=(0, 10))
        
        legend_inner = tk.Frame(legend_frame, bg='white')
        legend_inner.pack(fill='x', padx=15, pady=10)
        
        legend_title = tk.Label(legend_inner, text="Legend:", 
                               font=('Segoe UI', 9, 'bold'),
                               bg='white', fg=self.colors['text'])
        legend_title.pack(side='left', padx=(0, 15))
        
        # Notes indicator
        notes_indicator = tk.Label(legend_inner, text="📝 = Has Notes", 
                                  font=('Segoe UI', 9),
                                  bg='#d4edda', fg='#155724',
                                  padx=8, pady=2, relief=tk.FLAT)
        notes_indicator.pack(side='left', padx=(0, 10))
        
        # Search highlight indicator
        search_indicator = tk.Label(legend_inner, text="Search Match", 
                                   font=('Segoe UI', 9),
                                   bg='#fffacd', fg=self.colors['text'],
                                   padx=8, pady=2, relief=tk.FLAT)
        search_indicator.pack(side='left', padx=(0, 10))
        
        # Metric row indicator
        metric_indicator = tk.Label(legend_inner, text="Metric Row", 
                                   font=('Segoe UI', 9),
                                   bg='#e8f4f8', fg=self.colors['text'],
                                   padx=8, pady=2, relief=tk.FLAT)
        metric_indicator.pack(side='left')

        # --- Dashboard Treeview Frame ---
        dashboard_frame = tk.Frame(content_frame, bg='white', relief=tk.FLAT, bd=0)
        dashboard_frame.pack(expand=True, fill='both')
        
        # Treeview with scrollbar
        dashboard_scroll = ttk.Scrollbar(dashboard_frame)
        dashboard_scroll.pack(side='right', fill='y')
        
        # Configure treeview style
        style.configure('Custom.Treeview',
                       background='white',
                       foreground=self.colors['text'],
                       fieldbackground='white',
                       font=('Segoe UI', 9))
        style.configure('Custom.Treeview.Heading',
                       background=self.colors['primary'],
                       foreground='white',
                       font=('Segoe UI', 10, 'bold'))
        style.map('Custom.Treeview.Heading',
                 background=[('active', '#1a252f')])
        
        self.dashboard_tree = ttk.Treeview(dashboard_frame, 
                                          yscrollcommand=dashboard_scroll.set,
                                          style='Custom.Treeview')
        dashboard_scroll.config(command=self.dashboard_tree.yview)
        self.dashboard_tree.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Treeview Color Configuration
        self.dashboard_tree.tag_configure('metric_row', background='#e8f4f8', 
                                         font=('Segoe UI', 9, 'bold'))
        self.dashboard_tree.tag_configure('highlight', background='#fffacd')
        self.dashboard_tree.tag_configure('has_notes', background='#d4edda', 
                                         foreground='#155724')  # Light green
        self.dashboard_tree.tag_configure('highlight_with_notes', background='#b8dab8')
        
        # Add context menu to dashboard tree
        def show_context_menu(event):
            # Select row under cursor
            item = self.dashboard_tree.identify_row(event.y)
            if item:
                self.dashboard_tree.selection_set(item)
                
                # Get equipment_id
                values = self.dashboard_tree.item(item)['values']
                if values and values[0] not in ['Failure Rate', 'Availability', 'Total Failures']:
                    # Remove note icon if present
                    eq_id = str(values[0]).replace('📝 ', '')
                    context_menu = tk.Menu(self, tearoff=0)
                    context_menu.add_command(
                        label=f"📝 View/Add Notes for {eq_id}", 
                        command=self.view_selected_equipment_notes
                    )
                    context_menu.post(event.x_root, event.y_root)
        
        self.dashboard_tree.bind("<Button-3>", show_context_menu)  # Right-click
        
        # Add keyboard shortcuts
        self.bind('<Control-i>', lambda e: run_in_thread(self._threaded_import_excel))
        self.bind('<Control-e>', lambda e: run_in_thread(self._threaded_export_dashboard))
        self.bind('<Control-r>', lambda e: run_in_thread(self._threaded_refresh_dashboard_table))
        self.bind('<F5>', lambda e: run_in_thread(self._threaded_refresh_dashboard_table))
        self.bind('<Control-f>', lambda e: self.search_entry.focus())
        self.bind('<Escape>', lambda e: self.clear_search())
        self.bind('<Control-n>', lambda e: self.view_selected_equipment_notes())

    # ----------------------------------------------------------------------
    # EQUIPMENT NOTES FUNCTIONALITY
    # ----------------------------------------------------------------------
    
    def get_equipment_with_notes(self):
        """Get set of equipment IDs that have notes"""
        try:
            with self.get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("""
                    SELECT DISTINCT equipment_id
                    FROM equipment_notes
                """)
                results = cursor.fetchall()
                return {row[0] for row in results}
        except:
            return set()
  
    def view_selected_equipment_notes(self):
        """Wrapper to view notes from menu/button"""
        self.show_equipment_notes()

    def show_equipment_notes(self, equipment_id=None):
        """Display and manage notes for selected equipment"""
        # Get equipment_id from selection if not provided
        if equipment_id is None:
            selected_items = self.dashboard_tree.selection()
            if not selected_items:
                messagebox.showwarning("No Selection", "Please select an equipment from the dashboard.")
                return
            
            values = self.dashboard_tree.item(selected_items[0])['values']
            if not values:
                return
                
            equipment_id = str(values[0]).replace('📝 ', '')  # Remove note icon if present
            
            # Skip if it's a metric row
            if equipment_id in ['Failure Rate', 'Availability', 'Total Failures']:
                messagebox.showinfo("Invalid Selection", "Please select an equipment row, not a metric row.")
                return
        
        # Create notes window
        notes_window = tk.Toplevel(self)
        notes_window.title(f"Equipment Notes - {equipment_id}")
        notes_window.geometry("700x600")
        notes_window.configure(bg=self.colors['bg'])
        
        # Header
        header_frame = tk.Frame(notes_window, bg='white', relief=tk.FLAT, bd=0)
        header_frame.pack(fill='x', padx=0, pady=0)
        
        header_inner = tk.Frame(header_frame, bg='white')
        header_inner.pack(fill='x', padx=20, pady=15)
        
        title_label = tk.Label(header_inner, text=f"📝 Notes for: {equipment_id}", 
                              font=('Segoe UI', 14, 'bold'), 
                              bg='white', fg=self.colors['primary'])
        title_label.pack(side="left")
        
        separator = tk.Frame(notes_window, height=2, bg='#ddd')
        separator.pack(fill='x')
        
        # Main content frame
        content_frame = tk.Frame(notes_window, bg=self.colors['bg'])
        content_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # Notes display frame
        display_frame = tk.Frame(content_frame, bg='white', relief=tk.FLAT, bd=0)
        display_frame.pack(fill='both', expand=True, pady=(0, 15))
        
        # Scrollbar and Listbox for notes
        notes_scroll = tk.Scrollbar(display_frame)
        notes_scroll.pack(side='right', fill='y')
        
        notes_listbox = tk.Listbox(display_frame, 
                                   yscrollcommand=notes_scroll.set,
                                   font=('Segoe UI', 9),
                                   selectmode=tk.SINGLE,
                                   bg='white',
                                   fg=self.colors['text'],
                                   selectbackground=self.colors['secondary'],
                                   selectforeground='white')
        notes_listbox.pack(fill='both', expand=True, padx=10, pady=10)
        notes_scroll.config(command=notes_listbox.yview)
        
        # Input frame
        input_frame = tk.Frame(content_frame, bg='white', relief=tk.FLAT, bd=0)
        input_frame.pack(fill='x', pady=(0, 15))
        
        input_inner = tk.Frame(input_frame, bg='white')
        input_inner.pack(fill='x', padx=15, pady=15)
        
        # Text entry
        entry_label = tk.Label(input_inner, text="Add New Note:", 
                              font=('Segoe UI', 10, 'bold'),
                              bg='white', fg=self.colors['text'])
        entry_label.pack(anchor='w', pady=(0, 5))
        
        note_text = tk.Text(input_inner, height=4, width=50,
                           font=('Segoe UI', 9),
                           wrap=tk.WORD,
                           bg='#f9f9f9',
                           relief=tk.SOLID,
                           bd=1)
        note_text.pack(fill='x', pady=(0, 10))
        
        # Author entry
        author_frame = tk.Frame(input_inner, bg='white')
        author_frame.pack(fill='x', pady=(0, 10))
        
        author_label = tk.Label(author_frame, text="Your Name:", 
                               font=('Segoe UI', 9),
                               bg='white', fg=self.colors['text'])
        author_label.pack(side='left', padx=(0, 10))
        
        author_entry = tk.Entry(author_frame, width=30,
                               font=('Segoe UI', 9),
                               bg='#f9f9f9',
                               relief=tk.SOLID,
                               bd=1)
        author_entry.pack(side='left')
        
        # Button frame
        button_frame = tk.Frame(input_inner, bg='white')
        button_frame.pack(fill='x')
        
        def load_notes():
            """Load and display notes for the equipment"""
            notes_listbox.delete(0, tk.END)
            
            try:
                with self.get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        SELECT note_text, created_date, created_by
                        FROM equipment_notes
                        WHERE equipment_id = ?
                        ORDER BY created_date DESC
                    """, (equipment_id,))
                    
                    notes = cursor.fetchall()
                    
                    if not notes:
                        notes_listbox.insert(tk.END, "No notes available for this equipment.")
                        notes_listbox.itemconfig(0, {'fg': '#999'})
                    else:
                        for note in notes:
                            note_text_val, created_date, created_by = note
                            author_info = f" by {created_by}" if created_by else ""
                            
                            # Format date
                            try:
                                dt = datetime.strptime(created_date, DATETIME_FORMAT)
                                date_str = dt.strftime('%Y-%m-%d %H:%M')
                            except:
                                date_str = created_date
                            
                            display_text = f"[{date_str}]{author_info}: {note_text_val}"
                            notes_listbox.insert(tk.END, display_text)
                            
            except Exception as err:
                log_message(f"Error loading notes: {err}", "ERROR")
                messagebox.showerror("Error", f"Failed to load notes: {err}")
        
        def add_note():
            """Add a new note to the database"""
            note_content = note_text.get("1.0", tk.END).strip()
            author = author_entry.get().strip()
            
            if not note_content:
                messagebox.showwarning("Empty Note", "Please enter a note before adding.")
                return
            
            try:
                with self.get_db_connection() as conn:
                    cursor = conn.cursor()
                    cursor.execute("""
                        INSERT INTO equipment_notes (equipment_id, note_text, created_by)
                        VALUES (?, ?, ?)
                    """, (equipment_id, note_content, author if author else None))
                    conn.commit()
                
                log_message(f"Note added for {equipment_id} by {author or 'Anonymous'}", "INFO")
                
                # Clear inputs
                note_text.delete("1.0", tk.END)
                author_entry.delete(0, tk.END)
                
                # Reload notes
                load_notes()
                
                # Refresh dashboard to update highlighting
                self.after(100, self.refresh_dashboard_table)
                
                messagebox.showinfo("Success", "Note added successfully!")
                
            except Exception as err:
                log_message(f"Error adding note: {err}", "ERROR")
                messagebox.showerror("Error", f"Failed to add note: {err}")
        
        def delete_note():
            """Delete the selected note"""
            selected = notes_listbox.curselection()
            if not selected:
                messagebox.showwarning("No Selection", "Please select a note to delete.")
                return
            
            confirm = messagebox.askyesno("Confirm Delete", 
                                         "Are you sure you want to delete this note?")
            if not confirm:
                return
            
            try:
                # Get the note text from selection
                note_display = notes_listbox.get(selected[0])
                
                if "No notes available" in note_display:
                    return
                
                # Parse the note to get the actual text
                # Format: "[date] by author: text" or "[date]: text"
                if "]: " in note_display:
                    note_content = note_display.split("]: ", 1)[1]
                else:
                    note_content = note_display
                
                with self.get_db_connection() as conn:
                    cursor = conn.cursor()
                    # Delete the most recent matching note
                    cursor.execute("""
                        DELETE FROM equipment_notes
                        WHERE id = (
                            SELECT id FROM equipment_notes
                            WHERE equipment_id = ? AND note_text = ?
                            ORDER BY created_date DESC
                            LIMIT 1
                        )
                    """, (equipment_id, note_content))
                    conn.commit()
                
                log_message(f"Note deleted for {equipment_id}", "INFO")
                load_notes()
                
                # Refresh dashboard to update highlighting
                self.after(100, self.refresh_dashboard_table)
                
                messagebox.showinfo("Success", "Note deleted successfully!")
                
            except Exception as err:
                log_message(f"Error deleting note: {err}", "ERROR")
                messagebox.showerror("Error", f"Failed to delete note: {err}")
        
        # Add buttons
        btn_add = tk.Button(button_frame, text="➕ Add Note",
                           command=add_note,
                           bg=self.colors['success'], fg='white',
                           activebackground='#229954',
                           font=('Segoe UI', 9), relief=tk.FLAT, bd=0,
                           padx=15, pady=8, cursor='hand2')
        btn_add.pack(side='left', padx=(0, 10))
        
        btn_delete = tk.Button(button_frame, text="🗑️ Delete Selected",
                              command=delete_note,
                              bg=self.colors['danger'], fg='white',
                              activebackground='#c0392b',
                              font=('Segoe UI', 9), relief=tk.FLAT, bd=0,
                              padx=15, pady=8, cursor='hand2')
        btn_delete.pack(side='left', padx=(0, 10))
        
        btn_refresh = tk.Button(button_frame, text="🔄 Refresh",
                               command=load_notes,
                               bg=self.colors['info'], fg='white',
                               activebackground='#2980b9',
                               font=('Segoe UI', 9), relief=tk.FLAT, bd=0,
                               padx=15, pady=8, cursor='hand2')
        btn_refresh.pack(side='left')
        
        # Load initial notes
        load_notes()
        
        # Focus on text entry
        note_text.focus()

    # ----------------------------------------------------------------------
    # SEARCH AND STATISTICS
    # ----------------------------------------------------------------------
    
    def clear_search(self):
        """Clear the search box"""
        self.search_var.set("")
        self.search_entry.focus()
    
    def on_search_change(self, *args):
        """Filter dashboard tree based on search text"""
        search_text = self.search_var.get().lower().strip()
        
        if self._full_data is None or self._full_data.empty:
            return
        
        # Clear current tree
        for item in self.dashboard_tree.get_children():
            self.dashboard_tree.delete(item)
        
        if not search_text:
            self._populate_tree(self._full_data, highlight=False)
            return
        
        # Filter data - need to clean equipment_id for comparison
        def clean_eq_id(eq_id):
            return str(eq_id).replace('📝 ', '').lower()
        
        mask = (
            self._full_data['equipment_id'].apply(clean_eq_id).str.contains(search_text, na=False) |
            self._full_data['equipment_id'].isin(['Failure Rate', 'Availability', 'Total Failures'])
        )
        
        filtered_df = self._full_data[mask]
        
        if filtered_df.empty:
            self.dashboard_tree["columns"] = ("No Results",)
            self.dashboard_tree["show"] = "headings"
            self.dashboard_tree.heading("No Results", text=f"No equipment found matching '{search_text}'")
            self.dashboard_tree.column("No Results", width=500)
        else:
            self._populate_tree(filtered_df, highlight=True)
    
    def _populate_tree(self, df, highlight=False):
        """Helper method to populate the tree with data"""
        if df.empty:
            return
        
        # Get equipment with notes
        equipment_with_notes = self.get_equipment_with_notes()
        
        # Ensure columns are set
        self.dashboard_tree["columns"] = list(df.columns)
        self.dashboard_tree["show"] = "headings"
        
        for col in df.columns:
            self.dashboard_tree.heading(
                col, 
                text=str(col), 
                command=lambda c=col: self.sort_dashboard_column(c, False)
            )
            self.dashboard_tree.column(col, width=100)
        
        for _, row in df.iterrows():
            equipment_id = str(row['equipment_id'])
            has_notes = equipment_id in equipment_with_notes
            
            # Determine tag based on row type and notes status
            if equipment_id in ['Failure Rate', 'Availability', 'Total Failures']:
                tag = 'metric_row'
            elif highlight and self.search_var.get():
                # If searching and has notes, combine both highlights
                tag = 'highlight_with_notes' if has_notes else 'highlight'
            elif has_notes:
                tag = 'has_notes'
            else:
                tag = ''
            
            # Prepare values
            values = list(row)
            
            # Add visual indicator in the equipment_id column if has notes
            if has_notes and equipment_id not in ['Failure Rate', 'Availability', 'Total Failures']:
                values[0] = f"📝 {equipment_id}"
            
            # Insert with tag
            self.dashboard_tree.insert('', 'end', values=values, tags=(tag,))
    
    def show_database_stats(self):
        """Show database statistics"""
        if not self.selected_shaft.get():
            messagebox.showwarning("No Site Selected", "Please select a site first.")
            return
        
        try:
            with self.get_db_connection() as conn:
                cursor = conn.cursor()
                
                # Get total records
                cursor.execute("SELECT COUNT(*) FROM sensor_tests")
                total_records = cursor.fetchone()[0]
                
                # Get date range
                cursor.execute("""
                    SELECT MIN(DATE(time_tested)), MAX(DATE(time_tested)) 
                    FROM sensor_tests
                """)
                date_range = cursor.fetchone()
                
                # Get unique equipment count
                cursor.execute("SELECT COUNT(DISTINCT equipment_id) FROM sensor_tests")
                unique_equipment = cursor.fetchone()[0]
                
                # Get equipment with notes count
                cursor.execute("SELECT COUNT(DISTINCT equipment_id) FROM equipment_notes")
                equipment_with_notes = cursor.fetchone()[0]
                
                # Get total notes count
                cursor.execute("SELECT COUNT(*) FROM equipment_notes")
                total_notes = cursor.fetchone()[0]
                
                # Get pass/fail counts
                cursor.execute("""
                    SELECT outcome, COUNT(*) 
                    FROM sensor_tests 
                    GROUP BY outcome
                """)
                outcome_counts = dict(cursor.fetchall())
                
                # Get database file size
                db_path = self.get_db_name()
                db_size = os.path.getsize(db_path) / (1024 * 1024)  # MB
                
                # Get top 5 equipment with most test DAYS (using dashboard logic)
                query = """
                    SELECT equipment_id, time_tested
                    FROM sensor_tests
                """
                df = pd.read_sql(query, conn)
                
                if not df.empty:
                    # Apply same logic as dashboard
                    df['time_tested'] = pd.to_datetime(df['time_tested'])
                    df['minute'] = df['time_tested'].dt.floor('min')
                    df['date'] = df['time_tested'].dt.date
                    
                    # Group by equipment and minute first
                    grouped = df.groupby(['equipment_id', 'minute']).size().reset_index(name='count')
                    grouped['date'] = pd.to_datetime(grouped['minute']).dt.date
                    
                    # Count unique test days per equipment
                    daily_tests = grouped.groupby(['equipment_id', 'date']).size().reset_index(name='daily_count')
                    equipment_test_days = daily_tests.groupby('equipment_id').size().reset_index(name='test_days')
                    
                    # Sort and get top 5
                    top_equipment = equipment_test_days.sort_values(by='test_days', ascending=False).head(5)
                    top_equipment_list = list(top_equipment.itertuples(index=False, name=None))
                else:
                    top_equipment_list = []
                
                stats_msg = (
                    f"╔═══════════════════════════════════════╗\n"
                    f"  DATABASE STATISTICS\n"
                    f"  Site: {self.selected_shaft.get()}\n"
                    f"╚═══════════════════════════════════════╝\n\n"
                    f"📊 GENERAL STATISTICS\n"
                    f"─────────────────────────────────────────\n"
                    f"  Total Records: {total_records:,}\n"
                    f"  Unique Equipment: {unique_equipment:,}\n"
                    f"  Date Range: {date_range[0] or 'N/A'} to {date_range[1] or 'N/A'}\n"
                    f"  Database Size: {db_size:.2f} MB\n\n"
                    f"📝 NOTES STATISTICS\n"
                    f"─────────────────────────────────────────\n"
                    f"  Equipment with Notes: {equipment_with_notes:,}\n"
                    f"  Total Notes: {total_notes:,}\n\n"
                    f"✅ OUTCOMES\n"
                    f"─────────────────────────────────────────\n"
                )
                
                for outcome, count in outcome_counts.items():
                    percentage = (count / total_records * 100) if total_records > 0 else 0
                    stats_msg += f"  {outcome.capitalize()}: {count:,} ({percentage:.1f}%)\n"
                
                if top_equipment_list:
                    stats_msg += (
                        f"\n🏆 TOP 5 MOST TESTED EQUIPMENT (by Days)\n"
                        f"─────────────────────────────────────────\n"
                    )
                    for idx, (eq_id, test_days) in enumerate(top_equipment_list, 1):
                        stats_msg += f"  {idx}. {eq_id}: {test_days:,} days tested\n"
                
                messagebox.showinfo("Database Statistics", stats_msg)
                log_message(f"Viewed stats for {self.selected_shaft.get()}: {total_records:,} records", "INFO")
                
        except Exception as err:
            messagebox.showerror("Statistics Error", str(err))
            log_message(f"Error showing statistics: {err}", "ERROR")

    # ----------------------------------------------------------------------
    # UI and THREAD MANAGEMENT
    # ----------------------------------------------------------------------
    
    def _set_status(self, text, color="black"):
        """Updates the status label on the main thread"""
        self.after(0, lambda: self.status_label.config(text=text, fg=color))

    def show_progress(self, show=True):
        """Show or hide progress bar"""
        if show:
            self.progress_bar.pack(fill='x', padx=20, pady=(0, 10))
            self.progress_var.set(0)
        else:
            self.progress_bar.pack_forget()
            self.progress_var.set(0)

    def _threaded_import_excel(self):
        """Threaded Excel import"""
        self._set_status("Importing files...", "blue")
        self.after(0, lambda: self.show_progress(True))
        
        if not self.selected_shaft.get():
            self.after(0, lambda: messagebox.showwarning("No Site Selected", 
                                                         "Please select a site before importing data."))
            self._set_status("Ready", "green")
            self.after(0, lambda: self.show_progress(False))
            return
            
        self.import_excel()
        self._set_status("Import complete.", "green")
        self.after(0, lambda: self.show_progress(False))
        self._threaded_refresh_dashboard_table()

    def _threaded_export_dashboard(self):
        """Threaded dashboard export"""
        self._set_status("Exporting dashboard...", "blue")
        if not self.selected_shaft.get():
            self.after(0, lambda: messagebox.showwarning("No Site Selected", 
                                                         "Please select a site before exporting data."))
            self._set_status("Ready", "green")
            return
            
        self.export_dashboard()
        self._set_status("Export complete.", "green")

    def _threaded_export_daily_fail_count_report(self):
        """Threaded fail count report export"""
        self._set_status("Exporting fail count report...", "blue")
        if not self.selected_shaft.get():
            self.after(0, lambda: messagebox.showwarning("No Site Selected", 
                                                         "Please select a site before exporting data."))
            self._set_status("Ready", "green")
            return
            
        self.export_daily_fail_count_report()
        self._set_status("Export complete.", "green")

    def _threaded_refresh_dashboard_table(self):
        """Threaded dashboard refresh"""
        self._set_status("Loading dashboard...", "blue")
        if self.selected_shaft.get():
            self.refresh_dashboard_table()
        self._set_status("Ready", "green")
        
    def _threaded_show_failure_trend(self):
        """Threaded failure trend chart generation"""
        self._set_status("Generating failure trend chart...", "blue")
        if not self.selected_shaft.get():
            self.after(0, lambda: messagebox.showwarning("No Site Selected", 
                                                         "Please select a site."))
            self._set_status("Ready", "green")
            return
        
        df = self._get_daily_metrics_df()
        self.after(0, lambda: self._display_failure_chart(df))
        self._set_status("Ready", "green")

    def _threaded_show_consolidated_failure_trend(self):
        """Threaded consolidated failure trend chart generation"""
        self._set_status("Generating consolidated failure trend chart...", "blue")
        df = self._get_consolidated_daily_metrics()
        self.after(0, lambda: self._display_consolidated_failure_chart(df))
        self._set_status("Ready", "green")

    # ----------------------------------------------------------------------
    # CORE DATA LOGIC
    # ----------------------------------------------------------------------

    def get_db_name(self):
        """Returns the full path to the currently selected SQLite database"""
        selected_key = self.selected_shaft.get()
        if not selected_key or selected_key not in self.shaft_databases_cache:
            raise ValueError("No valid site selected for database operation.")
            
        db_filename = self.shaft_databases_cache[selected_key]
        return os.path.join(APP_DATA_DIR, db_filename)

    @contextlib.contextmanager
    def get_db_connection(self):
        """Context manager for database connections with better error handling"""
        try:
            db_name = self.get_db_name()
        except ValueError as err:
            raise ValueError(f"Database connection failed: {err}")
        
        conn = None
        try:
            conn = sqlite3.connect(db_name, timeout=10.0)
            conn.row_factory = sqlite3.Row
            yield conn
        except sqlite3.Error as err:
            if conn:
                conn.rollback()
            raise
        finally:
            if conn:
                conn.close()

    def validate_dataframe(self, df, required_columns):
        """Validates that a DataFrame has the required columns"""
        missing = [col for col in required_columns if col not in df.columns]
        return len(missing) == 0, missing

    def init_db(self, db_name=None):
        """Initializes the database structure"""
        if db_name is None:
            try:
                db_name = self.get_db_name()
            except ValueError: 
                log_message("DB Initialization skipped: No site selected", "WARNING")
                return
        
        try:
            conn = sqlite3.connect(db_name, timeout=10.0)
            with conn:
                cursor = conn.cursor()
                
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS sensor_tests (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        equipment_id TEXT,
                        serial TEXT,
                        equipment_type TEXT,
                        employee_id TEXT,
                        technician_name TEXT,
                        section TEXT,
                        shift TEXT,
                        time_tested DATETIME,
                        gas_type TEXT,
                        measured_value REAL,
                        outcome TEXT
                    );
                """)
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_time_tested ON sensor_tests(time_tested);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_equipment_id ON sensor_tests(equipment_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_outcome ON sensor_tests(outcome);")
                
                # Create equipment notes table
                cursor.execute("""
                    CREATE TABLE IF NOT EXISTS equipment_notes (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        equipment_id TEXT NOT NULL,
                        note_text TEXT NOT NULL,
                        created_date DATETIME DEFAULT CURRENT_TIMESTAMP,
                        created_by TEXT,
                        FOREIGN KEY (equipment_id) REFERENCES sensor_tests(equipment_id)
                    );
                """)
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_notes_equipment_id ON equipment_notes(equipment_id);")
                cursor.execute("CREATE INDEX IF NOT EXISTS idx_notes_date ON equipment_notes(created_date);")
                
            log_message(f"Database initialized: {os.path.basename(db_name)}", "INFO")
        except Exception as err:
            log_message(f"Database initialization error: {err}", "ERROR")
        finally:
            if 'conn' in locals() and conn:
                conn.close()

    def add_new_site(self):
        """Add a new site to the configuration"""
        new_site = simpledialog.askstring("Add New Site", "Enter new site name:", parent=self)
        if not new_site:
            return
        new_site = new_site.strip()
        if new_site in self.shaft_databases_cache:
            messagebox.showinfo("Site Exists", f"The site '{new_site}' already exists.")
            return
        
        db_name = f"sentinel_{new_site.lower().replace(' ', '_')}.db" 
        self.shaft_databases_cache[new_site] = db_name
        save_shaft_databases(self.shaft_databases_cache)
        self.init_db(os.path.join(APP_DATA_DIR, db_name))
        self.shaft_dropdown['values'] = list(self.shaft_databases_cache.keys())
        self.selected_shaft.set(new_site)
        messagebox.showinfo("Success", f"Site '{new_site}' added successfully.")
        log_message(f"New site added: {new_site}", "INFO")
        self.refresh_dashboard_table()

    def remove_site(self):
        """Remove a site from the configuration"""
        selected = self.selected_shaft.get()
        if not selected:
            messagebox.showwarning("No Site Selected", "Please select a site to remove.")
            return

        confirm_data_delete = messagebox.askyesno(
            "Confirm Data Deletion", 
            f"Do you also want to PERMANENTLY DELETE the database file for '{selected}'?\n"
            "(Choosing 'No' only removes it from the list.)"
        )
        
        confirm_site_removal = messagebox.askyesno(
            "Confirm Site Removal", 
            f"Are you sure you want to remove the site '{selected}' from the list?"
        )
        
        if not confirm_site_removal:
            return

        db_file = self.shaft_databases_cache.get(selected)
        if db_file:
            del self.shaft_databases_cache[selected]
            save_shaft_databases(self.shaft_databases_cache)

            if confirm_data_delete:
                full_db_path = os.path.join(APP_DATA_DIR, db_file) 
                if os.path.exists(full_db_path):
                    try:
                        os.remove(full_db_path)
                        messagebox.showinfo("File Deleted", 
                                          f"Database file for '{selected}' was deleted.")
                        log_message(f"Database deleted: {db_file}", "INFO")
                    except Exception as err:
                        messagebox.showerror("File Deletion Error", 
                                           f"Could not delete database file: {err}")
                        log_message(f"Failed to delete database: {err}", "ERROR")

            self.shaft_dropdown['values'] = list(self.shaft_databases_cache.keys())
            if self.shaft_databases_cache:
                self.selected_shaft.set(list(self.shaft_databases_cache.keys())[0])
            else:
                self.selected_shaft.set("")
                
            self.refresh_dashboard_table() 
            messagebox.showinfo("Site Removed", 
                              f"Site '{selected}' has been removed from the list.")
            log_message(f"Site removed: {selected}", "INFO")

    def reset_database(self):
        """Reset the database for the selected site"""
        selected = self.selected_shaft.get()
        if not selected:
            messagebox.showwarning("No Site Selected", 
                                  "Please select a site before resetting its database.")
            return

        confirm = messagebox.askyesno(
            "Confirm Database Reset", 
            f"Are you sure you want to PERMANENTLY DELETE ALL TEST DATA from the "
            f"database for '{selected}'?"
        )
        if not confirm:
            return

        try:
            with self.get_db_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM sensor_tests;")
                cursor.execute("DELETE FROM equipment_notes;")
                conn.commit()
            
            messagebox.showinfo("Database Reset", 
                              f"All data for site '{selected}' has been successfully deleted.")
            log_message(f"Database reset: {selected}", "INFO")
            self.refresh_dashboard_table() 
        
        except Exception as err:
            messagebox.showerror("Database Reset Error", str(err))
            log_message(f"Database reset error: {err}", "ERROR")

    def import_excel(self):
        """Import Excel files into the database"""
        try:
            db_name = self.get_db_name()
        except ValueError as err:
            self.after(0, lambda e=err: messagebox.showwarning("Error", str(e)))
            return

        file_paths = filedialog.askopenfilenames(
            filetypes=[("Excel files", "*.xls *.xlsx")]
        )
        if not file_paths:
            self.after(0, lambda: self.show_progress(False))
            return
        
        required_cols = [
            'equipment_id', 'serial', 'equipment_type', 'employee_id',
            'technician_name', 'section', 'shift', 'time_tested',
            'gas_type', 'measured_value', 'outcome'
        ]
            
        total_files = len(file_paths)
        log_message(f"Starting import of {total_files} file(s)", "INFO")
        
        try:
            with sqlite3.connect(db_name) as conn:
                cursor = conn.cursor()
                
                for file_idx, file_path in enumerate(file_paths, 1):
                    progress = (file_idx / total_files) * 100
                    self.after(0, lambda p=progress: self.progress_var.set(p))
                    
                    file_name = os.path.basename(file_path)
                    
                    self.after(0, lambda fn=file_name, idx=file_idx, tot=total_files: 
                        self._set_status(f"Processing {fn} ({idx}/{tot})...", "blue"))
                    
                    if file_path.lower().endswith('.xls'):
                        engine_type = 'xlrd' 
                    else:
                        engine_type = 'openpyxl' 
                    
                    try:
                        df = pd.read_excel(file_path, engine=engine_type)

                        df = df.rename(columns={
                            'Equipment Id': 'equipment_id', 
                            'Serial': 'serial', 
                            'Equipment Type': 'equipment_type', 
                            'Employee': 'employee_id', 
                            'Name': 'technician_name', 
                            'Section': 'section', 
                            'Shift': 'shift', 
                            'Time Tested': 'time_tested', 
                            'Gas Type': 'gas_type', 
                            'Measured Value': 'measured_value', 
                            'Outcome': 'outcome'
                        })
                        
                        is_valid, missing = self.validate_dataframe(df, required_cols)
                        if not is_valid:
                            error_msg = f"Missing required columns: {', '.join(missing)}"
                            raise ValueError(error_msg)
                        
                        rows_before = len(df)
                        
                        df['time_tested'] = pd.to_datetime(df['time_tested'], errors='coerce')
                        temp_serials = pd.to_numeric(df['time_tested'].copy(), errors='coerce')
                        valid_serial_mask = (temp_serials > 1) & (temp_serials < 90000)
                        serial_dates = pd.to_datetime(
                            temp_serials[valid_serial_mask], 
                            unit='D', 
                            origin=EXCEL_DATE_ORIGIN
                        )
                        df['time_tested'].fillna(serial_dates, inplace=True)

                        df.dropna(subset=['time_tested'], inplace=True)
                        rows_after = len(df)
                        
                        if rows_before > rows_after:
                            dropped_count = rows_before - rows_after
                            log_message(f"{dropped_count} rows dropped from '{file_name}' due to invalid dates", "WARNING")
                            self.after(0, lambda fn=file_name, dc=dropped_count: 
                                messagebox.showwarning(
                                    "Data Validation Warning",
                                    f"File: {fn}\n{dc} rows were skipped due to invalid date formats."
                                ))

                        df['time_tested'] = pd.to_datetime(df['time_tested']).dt.strftime(DATETIME_FORMAT)
                        df.drop_duplicates(subset=['equipment_id', 'time_tested'], inplace=True)
                        
                        insert_cols = ['equipment_id', 'serial', 'equipment_type', 
                                       'employee_id', 'technician_name', 'section', 
                                       'shift', 'time_tested', 'gas_type', 
                                       'measured_value', 'outcome']
                                       
                        df_to_insert = df[insert_cols].copy()
                        records = df_to_insert.values.tolist()
                        
                        cursor.executemany("""
                            INSERT INTO sensor_tests (
                                equipment_id, serial, equipment_type, employee_id, 
                                technician_name, section, shift, time_tested, gas_type, 
                                measured_value, outcome
                            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """, records)
                        
                        log_message(f"Imported {len(records)} records from {file_name}", "INFO")
                        
                    except Exception as file_err:
                        error_message = str(file_err)
                        log_message(f"Error processing {file_name}: {error_message}", "ERROR")
                        
                        self.after(0, lambda fn=file_name, msg=error_message: 
                            messagebox.showwarning(
                                "File Import Error", 
                                f"Skipping file {fn} due to error:\n{msg}"
                            ))
                        continue
                    
                conn.commit()
            
        except Exception as db_error:
            log_message(f"Database import error: {db_error}", "ERROR")
            self.after(0, lambda err=db_error: 
                messagebox.showerror("Import Error (Database)", str(err)))

    def sort_dashboard_column(self, col, reverse):
        """Sort dashboard columns with intelligent type detection"""
        data = [(self.dashboard_tree.set(k, col), k) 
                for k in self.dashboard_tree.get_children('')]
        
        try:
            def sort_key(t):
                s = str(t[0]).replace('%', '').replace(',', '').replace('📝 ', '').strip()
                try:
                    return float(s)
                except ValueError:
                    return t[0] 
                    
            data.sort(key=sort_key, reverse=reverse)
        except Exception:
            data.sort(key=lambda t: t[0], reverse=reverse)
            
        for index, (_, k) in enumerate(data):
            self.dashboard_tree.move(k, '', index)
        
        self.dashboard_tree.heading(
            col, 
            command=lambda: self.sort_dashboard_column(col, not reverse)
        )

    def calculate_metrics(self, pivot_df):
        """Calculates Total Failures, Failure Rate, and Availability"""
        daily_failures = (pivot_df.iloc[:, 1:-1] == 'fail').sum(axis=0)
        daily_tests = (pivot_df.iloc[:, 1:-1].isin(['pass', 'fail'])).sum(axis=0)
        total_failures = daily_failures.apply(lambda x: int(x)) 
        failure_rate = (daily_failures / daily_tests).fillna(0).apply(
            lambda x: f"{x:.2%}" if pd.notna(x) else '0.00%'
        )
        daily_passes = (pivot_df.iloc[:, 1:-1] == 'pass').sum(axis=0)
        availability = (daily_passes / daily_tests).fillna(1).apply(
            lambda x: f"{x:.2%}" if pd.notna(x) else '100.00%'
        )

        metrics_data = {
            pivot_df.columns[0]: ['Total Failures', 'Failure Rate', 'Availability'], 
            **{col: [
                total_failures.loc[col],  
                failure_rate.loc[col], 
                availability.loc[col]
            ] for col in daily_tests.index}
        }
        
        metrics_data[pivot_df.columns[-1]] = [0, '', ''] 
        metrics_df = pd.DataFrame(metrics_data).set_index(
            pivot_df.columns[0]
        ).reset_index()

        return metrics_df

    def _get_cache_key(self):
        """Generate a cache key based on current filter settings"""
        return (
            self.selected_shaft.get(),
            self.from_date.get_date(),
            self.to_date.get_date()
        )
        
    def _get_dashboard_df(self, use_cache=True):
        """Retrieves, processes, and calculates metrics for the dashboard view"""
        current_key = self._get_cache_key()
        
        if use_cache and current_key == self._cache_key and self._dashboard_cache is not None:
            return self._dashboard_cache.copy()

        start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
        end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')
        
        try:
            with self.get_db_connection() as conn:
                query = """
                SELECT equipment_id, time_tested, outcome
                FROM sensor_tests
                WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                """
                df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])
        except ValueError:
            return pd.DataFrame()
        except Exception as err:
            log_message(f"Database query error: {err}", "ERROR")
            return pd.DataFrame()
        
        if df.empty:
            return pd.DataFrame()

        df['date'] = pd.to_datetime(df['time_tested']).dt.date
        df['minute'] = pd.to_datetime(df['time_tested']).dt.floor('min')
        
        grouped = df.groupby(['equipment_id', 'minute']).agg({
            'outcome': lambda x: 'fail' if 'fail' in x.values else 'pass'
        }).reset_index()
        grouped['date'] = grouped['minute'].dt.date
        
        pivot_df = grouped.pivot_table(
            index='equipment_id', 
            columns='date', 
            values='outcome', 
            aggfunc='first'
        )
        pivot_df = pivot_df.fillna("-")
        pivot_df['Failed'] = (pivot_df == 'fail').sum(axis=1)
        pivot_df.reset_index(inplace=True)
        
        metrics_df = self.calculate_metrics(pivot_df.copy())
        final_df = pd.concat([pivot_df, metrics_df], ignore_index=True)
        
        self._cache_key = current_key
        self._dashboard_cache = final_df.copy()
        
        return final_df

    def _get_daily_metrics_df(self):
        """Retrieves daily pass/fail counts and calculates daily metrics"""
        start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
        end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')
    
        try:
            with self.get_db_connection() as conn:
                query = """
                SELECT equipment_id, time_tested, outcome
                FROM sensor_tests
                WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                """
                df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])
        except ValueError:
            return pd.DataFrame()
        except Exception as err:
            log_message(f"Database query error: {err}", "ERROR")
            return pd.DataFrame()
    
        if df.empty:
            return pd.DataFrame() 

        df['date'] = pd.to_datetime(df['time_tested']).dt.date
        df['minute'] = pd.to_datetime(df['time_tested']).dt.floor('min')
        
        grouped = df.groupby(['equipment_id', 'minute']).agg({
            'outcome': lambda x: 'fail' if 'fail' in x.values else 'pass'
        }).reset_index()
        grouped['date'] = grouped['minute'].dt.date
        
        daily_equipment_outcome = grouped.groupby(['equipment_id', 'date'])['outcome'].agg(
            lambda x: 'fail' if 'fail' in x.values else 'pass'
        ).reset_index()
        
        daily_counts = daily_equipment_outcome.groupby('date')['outcome'].value_counts().unstack(
            fill_value=0
        ).reset_index()
        daily_counts.columns.name = None
    
        if 'pass' not in daily_counts.columns: 
            daily_counts['pass'] = 0
        if 'fail' not in daily_counts.columns: 
            daily_counts['fail'] = 0
        
        daily_counts['Total Tests'] = daily_counts['pass'] + daily_counts['fail']
        daily_counts['Failure Rate'] = daily_counts.apply(
            lambda row: row['fail'] / row['Total Tests'] if row['Total Tests'] > 0 else 0, 
            axis=1
        )
        daily_counts['date'] = pd.to_datetime(daily_counts['date'])
    
        return daily_counts

    def _get_consolidated_daily_metrics(self):
        """Retrieves daily metrics from ALL configured databases"""
        start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
        end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')
        
        all_site_data = []

        for site_name, db_filename in self.shaft_databases_cache.items():
            db_path = os.path.join(APP_DATA_DIR, db_filename)
            self.after(0, lambda sn=site_name: self._set_status(f"Aggregating data from {sn}...", "blue"))

            if not os.path.exists(db_path):
                log_message(f"Skipping site {site_name}: database file not found", "WARNING")
                continue
                
            try:
                with sqlite3.connect(db_path) as conn:
                    query = """
                    SELECT equipment_id, time_tested, outcome
                    FROM sensor_tests
                    WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                    """
                    df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])

            except Exception as err:
                log_message(f"Database query error for site {site_name}: {err}", "ERROR")
                continue
            
            if df.empty:
                continue

            df['date'] = pd.to_datetime(df['time_tested']).dt.date
            df['minute'] = pd.to_datetime(df['time_tested']).dt.floor('min')
            
            grouped = df.groupby(['equipment_id', 'minute']).agg({
                'outcome': lambda x: 'fail' if 'fail' in x.values else 'pass'
            }).reset_index()
            grouped['date'] = grouped['minute'].dt.date
            
            daily_counts = grouped.groupby('date')['outcome'].value_counts().unstack(
                fill_value=0
            ).reset_index()
            daily_counts.columns.name = None
            
            if 'pass' not in daily_counts.columns: 
                daily_counts['pass'] = 0
            if 'fail' not in daily_counts.columns: 
                daily_counts['fail'] = 0
            
            daily_counts['Total Tests'] = daily_counts['pass'] + daily_counts['fail']
            daily_counts['Failure Rate'] = daily_counts.apply(
                lambda row: row['fail'] / row['Total Tests'] if row['Total Tests'] > 0 else 0, 
                axis=1
            )
            daily_counts['date'] = pd.to_datetime(daily_counts['date'])
            daily_counts['Site'] = site_name

            all_site_data.append(daily_counts)

        if not all_site_data:
            return pd.DataFrame()
            
        return pd.concat(all_site_data, ignore_index=True)

    def _get_consolidated_fail_count_report(self):
        """Aggregates daily failure count for every equipment_id across ALL sites"""
        start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
        end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')
        
        all_site_reports = []
        
        for site_name, db_filename in self.shaft_databases_cache.items():
            db_path = os.path.join(APP_DATA_DIR, db_filename)
            self.after(0, lambda sn=site_name: self._set_status(f"Processing failures from {sn}...", "blue"))

            if not os.path.exists(db_path):
                continue
                
            try:
                with sqlite3.connect(db_path) as conn:
                    query = """
                    SELECT equipment_id, serial, time_tested, outcome
                    FROM sensor_tests
                    WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                    """
                    df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])

            except Exception as err:
                log_message(f"Database query error for site {site_name}: {err}", "ERROR")
                continue
            
            if df.empty:
                continue
                
            df['time_tested'] = pd.to_datetime(df['time_tested'])
            df['minute'] = df['time_tested'].dt.floor('min')
            
            grouped = df.groupby(['equipment_id', 'minute'])['outcome'].agg(
                lambda x: 'fail' if 'fail' in x.values else 'pass'
            ).reset_index()

            grouped['date'] = grouped['minute'].dt.date

            daily_outcome = grouped.groupby(['equipment_id', 'date'])['outcome'].agg(
                lambda x: 'fail' if 'fail' in x.values else 'pass'
            ).reset_index()

            fail_count_df = daily_outcome[
                daily_outcome['outcome'] == 'fail'
            ].groupby(['equipment_id'])['date'].nunique().reset_index(name='Failed Days Count') 

            serial_mode = df.groupby('equipment_id')['serial'].agg(
                lambda x: x.mode().iat[0] if not x.mode().empty else ''
            ).reset_index()
            
            fail_count_df = fail_count_df.merge(serial_mode, on='equipment_id', how='left')
            fail_count_df['Site'] = site_name
            all_site_reports.append(fail_count_df)

        if not all_site_reports:
            return pd.DataFrame()
            
        consolidated_df = pd.concat(all_site_reports, ignore_index=True)
        
        total_failures_by_equipment = consolidated_df.groupby('equipment_id').agg(
            {'Failed Days Count': 'sum', 'serial': 'first'}
        ).reset_index()
        total_failures_by_equipment = total_failures_by_equipment.rename(columns={
            'Failed Days Count': 'Total Failed Days (All Sites)'
        })
        
        site_list = consolidated_df.groupby('equipment_id')['Site'].apply(
            lambda x: ', '.join(sorted(x.unique()))
        ).reset_index(name='Sites Affected')

        consolidated_final = total_failures_by_equipment.merge(site_list, on='equipment_id')
        
        consolidated_final = consolidated_final[[
            'equipment_id', 
            'serial', 
            'Total Failed Days (All Sites)', 
            'Sites Affected'
        ]]
        
        consolidated_final = consolidated_final.sort_values(
            by='Total Failed Days (All Sites)', 
            ascending=False
        ).reset_index(drop=True)
        
        consolidated_final.index += 1
        consolidated_final.insert(0, 'Rank', consolidated_final.index)
        
        return consolidated_final

    def refresh_dashboard_table(self):
        """Refresh the dashboard table display"""
        self._dashboard_cache = None
        self._cache_key = None
        
        try:
            final_df = self._get_dashboard_df(use_cache=False)
            self._full_data = final_df.copy()
            self.after(0, lambda: self._update_treeview(final_df))
            
        except Exception as dash_error:
            log_message(f"Dashboard refresh error: {dash_error}", "ERROR")
            self.after(0, lambda err=dash_error: 
                messagebox.showerror("Dashboard Error", str(err)))

    def _update_treeview(self, final_df):
        """Handles the actual GUI update of the Treeview"""
        self.dashboard_tree.delete(*self.dashboard_tree.get_children())
        
        if final_df.empty:
            self.dashboard_tree["columns"] = ("No Data",)
            self.dashboard_tree["show"] = "headings"
            self.dashboard_tree.heading(
                "No Data", 
                text=f"No data available for site '{self.selected_shaft.get()}' "
                     "in the selected range."
            )
            self.dashboard_tree.column("No Data", width=500)
            return

        self._populate_tree(final_df, highlight=False)

    def _display_failure_chart(self, df):
        """Creates and displays the failure chart with toggleable view"""
        if df.empty:
            messagebox.showinfo("Chart View", 
                              "No data to display in the selected date range.")
            return

        try:
            chart_window = tk.Toplevel(self)
            chart_window.title("Failure Analysis Chart")
            chart_window.geometry("1000x650")
    
            site_name = self.selected_shaft.get()
            start_date = self.from_date.get_date().strftime('%Y-%m-%d')
            end_date = self.to_date.get_date().strftime('%Y-%m-%d')
            chart_window.title(
                f"Daily Failure Trend: {site_name} ({start_date} to {end_date})"
            )

            button_frame = tk.Frame(chart_window, bg=self.colors['bg'])
            button_frame.pack(fill='x', padx=10, pady=5)
    
            chart_frame = tk.Frame(chart_window)
            chart_frame.pack(fill='both', expand=True)
    
            view_state = tk.BooleanVar(value=True)

            def update_chart():
                plt.clf()
                
                fig, ax = plt.subplots(figsize=(10, 6))
                
                dates_num = mdates.date2num(pd.to_datetime(df['date']).dt.to_pydatetime())
                
                if view_state.get():
                    y_data = df['Failure Rate'] * 100
                    ylabel = "Failure Rate (%)"
                    y_max = 100
                    y_interval = 5
                    data_format = lambda x: f"{x:.1f}%"
                else:
                    y_data = df['fail']
                    ylabel = "Number of Failures"
                    y_max = max(df['fail']) * 1.2
                    y_interval = max(1, int(y_max / 20))
                    data_format = lambda x: str(int(x))
    
                ax.plot(dates_num, y_data, marker='o', linestyle='-', 
                       color='red', label='Failures')
    
                for x, y in zip(dates_num, y_data):
                    ax.annotate(data_format(y),
                              (x, y),
                              xytext=(0, 10),
                              textcoords='offset points',
                              ha='center',
                              va='bottom',
                              fontsize=8)

                ax.set_title(f"Daily Failure {'Rate' if view_state.get() else 'Count'} for {site_name}", 
                            fontsize=14)
                ax.set_xlabel("Date", fontsize=12)
                ax.set_ylabel(ylabel, fontsize=12)
                ax.yaxis.set_major_locator(ticker.MultipleLocator(y_interval))
                ax.grid(True, linestyle='--', alpha=0.7)
                ax.legend()
                ax.set_ylim(0, y_max)
    
                ax.xaxis.set_major_locator(mdates.DayLocator(interval=2))
                ax.xaxis.set_minor_locator(mdates.DayLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
                fig.autofmt_xdate(rotation=45)
    
                if hasattr(chart_window, 'canvas'):
                    chart_window.canvas.get_tk_widget().destroy()
                
                canvas = FigureCanvasTkAgg(fig, master=chart_frame)
                canvas_widget = canvas.get_tk_widget()
                canvas_widget.pack(fill=tk.BOTH, expand=True)
                canvas.draw()
                
                chart_window.canvas = canvas
                chart_window.fig = fig

            def toggle_view():
                view_state.set(not view_state.get())
                update_chart()
    
            def on_closing():
                plt.close('all')
                chart_window.destroy()
    
            chart_window.protocol("WM_DELETE_WINDOW", on_closing)
    
            toggle_btn = tk.Button(
                button_frame,
                text="Toggle View (Percentage/Count)",
                command=toggle_view,
                bg=self.colors['secondary'],
                fg='white',
                activebackground='#2980b9',
                font=('Segoe UI', 9),
                relief=tk.FLAT,
                bd=0,
                padx=15,
                pady=8,
                cursor='hand2'
            )
            toggle_btn.pack(side='left')
    
            update_chart()
    
        except Exception as e:
            messagebox.showerror("Chart Error", f"Error creating chart: {str(e)}")
            log_message(f"Chart creation error: {str(e)}", "ERROR")
        
    def _display_consolidated_failure_chart(self, df):
        """Creates and displays the consolidated failure chart with toggleable view"""
        if df.empty:
            messagebox.showinfo("Chart View", 
                                "No consolidated data to display in the selected date range.")
            return

        try:
            chart_window = tk.Toplevel(self)
            chart_window.title("Consolidated Failure Analysis Chart")
            chart_window.geometry("1000x650")
            
            start_date = self.from_date.get_date().strftime('%Y-%m-%d')
            end_date = self.to_date.get_date().strftime('%Y-%m-%d')
            chart_window.title(
                f"Consolidated Daily Failure Trend ({start_date} to {end_date})"
            )

            # Create frames
            button_frame = tk.Frame(chart_window, bg=self.colors['bg'])
            button_frame.pack(fill='x', padx=10, pady=5)

            chart_frame = tk.Frame(chart_window)
            chart_frame.pack(fill='both', expand=True)

            # View state for toggle
            view_state = tk.BooleanVar(value=True)

            def update_chart():
                plt.clf()
                
                fig, ax = plt.subplots(figsize=(10, 6))
                
                for site_name, site_df in df.groupby('Site'):
                    site_df = site_df.sort_values(by='date')
                    dates_num = mdates.date2num(site_df['date'].dt.to_pydatetime())
                    
                    if view_state.get():
                        y_data = site_df['Failure Rate'] * 100
                        ylabel = "Failure Rate (%)"
                        y_max = 100
                        y_interval = 5
                    else:
                        y_data = site_df['fail']
                        ylabel = "Number of Failures"
                        max_fails = df.groupby('Site')['fail'].max().max()
                        y_max = max(max_fails * 1.2, 1)
                        y_interval = max(1, int(y_max / 20))
                    
                    ax.plot(dates_num, y_data, 
                           marker='.', linestyle='-', label=site_name, 
                           alpha=0.8, linewidth=2, markersize=6)
                
                ax.set_title(
                    f"Consolidated Daily Failure {'Rate' if view_state.get() else 'Count'} Trend Across All Sites", 
                    fontsize=14, fontweight='bold'
                )
                ax.set_xlabel("Date", fontsize=12)
                ax.set_ylabel(ylabel, fontsize=12)
                ax.yaxis.set_major_locator(ticker.MultipleLocator(y_interval))
                ax.grid(True, linestyle='--', alpha=0.5)
                ax.legend(title="Site", loc='upper left', framealpha=0.9)
                ax.set_ylim(0, y_max)
                
                num_days = len(df['date'].unique())
                if num_days <= 7:
                    interval = 1
                elif num_days <= 30:
                    interval = 2
                else:
                    interval = max(1, num_days // 15)
                    
                ax.xaxis.set_major_locator(mdates.DayLocator(interval=interval))
                ax.xaxis.set_minor_locator(mdates.DayLocator())
                ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m-%d'))
                fig.autofmt_xdate(rotation=45)
                fig.tight_layout()

                if hasattr(chart_window, 'canvas'):
                    chart_window.canvas.get_tk_widget().destroy()
                
                canvas = FigureCanvasTkAgg(fig, master=chart_frame)
                canvas_widget = canvas.get_tk_widget()
                canvas_widget.pack(fill=tk.BOTH, expand=True)
                canvas.draw()
                
                chart_window.canvas = canvas
                chart_window.fig = fig

            def toggle_view():
                view_state.set(not view_state.get())
                update_chart()

            def on_closing():
                if hasattr(chart_window, 'fig'):
                    plt.close(chart_window.fig)
                plt.close('all')
                chart_window.destroy()

            chart_window.protocol("WM_DELETE_WINDOW", on_closing)

            toggle_btn = tk.Button(
                button_frame,
                text="Toggle View (Percentage/Count)",
                command=toggle_view,
                bg=self.colors['secondary'],
                fg='white',
                activebackground='#2980b9',
                font=('Segoe UI', 9),
                relief=tk.FLAT,
                bd=0,
                padx=15,
                pady=8,
                cursor='hand2'
            )
            toggle_btn.pack(side='left')

            update_chart()

        except Exception as e:
            messagebox.showerror("Chart Error", f"Error creating chart: {str(e)}")
            log_message(f"Consolidated chart creation error: {str(e)}", "ERROR")

    def show_most_common_failure(self):
        """Show the most common failure analysis"""
        try:
            start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
            end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')
            
            with self.get_db_connection() as conn:
                query = """
                SELECT equipment_id, time_tested, outcome
                FROM sensor_tests
                WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                """
                df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])

            if df.empty:
                messagebox.showinfo("Analysis Result", 
                                  "No failures recorded in the selected date range.")
                return

            # Apply same logic as dashboard
            df['time_tested'] = pd.to_datetime(df['time_tested'])
            df['minute'] = df['time_tested'].dt.floor('min')
            df['date'] = df['time_tested'].dt.date
            
            # Group by equipment and minute first
            grouped = df.groupby(['equipment_id', 'minute'])['outcome'].agg(
                lambda x: 'fail' if 'fail' in x.values else 'pass'
            ).reset_index()
            grouped['date'] = grouped['minute'].dt.date
            
            # Get daily outcome per equipment
            daily_outcome = grouped.groupby(['equipment_id', 'date'])['outcome'].agg(
                lambda x: 'fail' if 'fail' in x.values else 'pass'
            ).reset_index()
            
            # Count failed days per equipment
            fail_count = daily_outcome[daily_outcome['outcome'] == 'fail'].groupby(
                'equipment_id'
            )['date'].nunique().reset_index(name='failed_days')
            
            # Sort by failed days
            fail_count = fail_count.sort_values(by='failed_days', ascending=False)
            
            if fail_count.empty:
                messagebox.showinfo("Analysis Result", 
                                  "No failures recorded in the selected date range.")
                return
            
            # Get top 10 most failed equipment
            top_10 = fail_count.head(10)
            
            # Build message
            result_msg = (
                f"╔═══════════════════════════════════════╗\n"
                f"  MOST COMMON FAILURES\n"
                f"  {start_date_str} to {end_date_str}\n"
                f"╚═══════════════════════════════════════╝\n\n"
                f"📊 SUMMARY\n"
                f"─────────────────────────────────────────\n"
                f"  Total Equipment with Failures: {len(fail_count):,}\n\n"
                f"🏆 TOP 10 MOST FAILED EQUIPMENT\n"
                f"─────────────────────────────────────────\n"
            )
            
            for idx, row in enumerate(top_10.itertuples(index=False), 1):
                equipment_id = row.equipment_id
                failed_days = row.failed_days
                result_msg += f"  {idx:2d}. {equipment_id}: {failed_days:,} failed days\n"
            
            messagebox.showinfo("Common Failures Analysis", result_msg)
            log_message(f"Common failures analysis: Top failure = {top_10.iloc[0]['equipment_id']} with {top_10.iloc[0]['failed_days']} failed days", "INFO")
            
        except ValueError as err:
            messagebox.showwarning("Error", str(err))
        except Exception as err:
            messagebox.showerror("Analysis Error", str(err))
            log_message(f"Error in common failures analysis: {err}", "ERROR")

    def export_dashboard(self):
        """Export dashboard with format options"""
        try:
            final_df = self._get_dashboard_df() 
            
            if final_df.empty:
                self.after(0, lambda: messagebox.showinfo(
                    "Export Error", 
                    "No data to export in the selected date range."
                ))
                return

            export_format = messagebox.askyesno(
                "Export Format",
                "Export as Excel?\n(No = Export as CSV)"
            )
            
            if export_format:
                file_types = [("Excel files", "*.xlsx")]
                default_ext = ".xlsx"
            else:
                file_types = [("CSV files", "*.csv")]
                default_ext = ".csv"
            
            export_path = filedialog.asksaveasfilename(
                defaultextension=default_ext,
                initialfile=f"Sentinel_Dashboard_{self.selected_shaft.get()}",
                filetypes=file_types
            )
            
            if export_path:
                # Clean equipment_id column (remove note icons) before export
                export_df = final_df.copy()
                export_df['equipment_id'] = export_df['equipment_id'].astype(str).str.replace('📝 ', '', regex=False)
                
                if export_format:
                    export_df.to_excel(export_path, index=False, engine='openpyxl')
                else:
                    export_df.to_csv(export_path, index=False, encoding='utf-8')
                    
                log_message(f"Dashboard exported to: {export_path}", "INFO")
                self.after(0, lambda: messagebox.showinfo(
                    "Export Success", 
                    f"Successfully exported data to:\n{export_path}"
                ))
                
        except Exception as export_err:
            log_message(f"Export error: {export_err}", "ERROR")
            self.after(0, lambda err=export_err: 
                messagebox.showerror("Export Error", str(err)))

    def export_daily_fail_count_report(self):
        """Export daily fail count report with separate tabs for each site and time-series chart"""
        try:
            start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
            end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')
            
            # Get individual site reports
            all_site_reports = {}
            all_site_daily_data = {}
            
            for site_name, db_filename in self.shaft_databases_cache.items():
                db_path = os.path.join(APP_DATA_DIR, db_filename)
                self.after(0, lambda sn=site_name: self._set_status(f"Processing {sn}...", "blue"))

                if not os.path.exists(db_path):
                    continue
                    
                try:
                    with sqlite3.connect(db_path) as conn:
                        query = """
                        SELECT equipment_id, serial, time_tested, outcome
                        FROM sensor_tests
                        WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                        """
                        site_df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])

                except Exception as err:
                    log_message(f"Database query error for site {site_name}: {err}", "ERROR")
                    continue
                
                if site_df.empty:
                    continue
                    
                # Process site data
                site_df['time_tested'] = pd.to_datetime(site_df['time_tested'])
                site_df['minute'] = site_df['time_tested'].dt.floor('min')
                
                grouped = site_df.groupby(['equipment_id', 'minute'])['outcome'].agg(
                    lambda x: 'fail' if 'fail' in x.values else 'pass'
                ).reset_index()

                grouped['date'] = grouped['minute'].dt.date

                daily_outcome = grouped.groupby(['equipment_id', 'date'])['outcome'].agg(
                    lambda x: 'fail' if 'fail' in x.values else 'pass'
                ).reset_index()

                fail_count_df = daily_outcome[
                    daily_outcome['outcome'] == 'fail'
                ].groupby(['equipment_id'])['date'].nunique().reset_index(name='Failed Days Count') 

                serial_mode = site_df.groupby('equipment_id')['serial'].agg(
                    lambda x: x.mode().iat[0] if not x.mode().empty else ''
                ).reset_index()
                
                site_fail_df = fail_count_df.merge(serial_mode, on='equipment_id', how='left')
                site_fail_df['Failed Days Count'] = site_fail_df['Failed Days Count'].astype(int)
                site_fail_df = site_fail_df[['equipment_id', 'serial', 'Failed Days Count']].sort_values(
                    by='Failed Days Count', 
                    ascending=False
                )
                
                all_site_reports[site_name] = site_fail_df
                
                # Get daily failure counts for chart
                daily_failures = daily_outcome[
                    daily_outcome['outcome'] == 'fail'
                ].groupby('date').size().reset_index(name='Failed Count')
                
                daily_failures['date'] = pd.to_datetime(daily_failures['date'])
                all_site_daily_data[site_name] = daily_failures
            
            if not all_site_reports:
                self.after(0, lambda: messagebox.showinfo(
                    "Export Error", 
                    "No data to export in the selected date range for any site."
                ))
                return

            export_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                initialfile=f"Daily_Fail_Count_Report_All_Sites_{start_date_str}_to_{end_date_str}",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if export_path:
                with pd.ExcelWriter(export_path, engine='openpyxl') as writer:
                    
                    # Write individual site reports
                    for site_name in sorted(all_site_reports.keys()):
                        site_report_df = all_site_reports[site_name]
                        if not site_report_df.empty:
                            sheet_name = site_name[:31]
                            site_report_df.to_excel(writer, sheet_name=sheet_name, index=False)
                            
                            if sheet_name in writer.sheets:
                                worksheet = writer.sheets[sheet_name]
                                try:
                                    serial_col_idx = site_report_df.columns.get_loc('serial') + 1 
                                    col_letter = chr(65 + serial_col_idx - 1) 
                                    worksheet.column_dimensions[col_letter].number_format = '@'
                                except (KeyError, IndexError):
                                    pass
                    
                    # Create consolidated summary with chart
                    if all_site_daily_data:
                        from openpyxl.chart import LineChart, Reference
                        
                        ws = writer.book.create_sheet('Consolidated Summary')
                        
                        all_dates = set()
                        for site_data in all_site_daily_data.values():
                            all_dates.update(site_data['date'].tolist())
                        
                        all_dates = sorted(list(all_dates))
                        
                        headers = ['Date'] + sorted(all_site_daily_data.keys())
                        ws.append(headers)
                        
                        for date in all_dates:
                            row_data = [date.strftime('%Y-%m-%d')]
                            
                            for site_name in sorted(all_site_daily_data.keys()):
                                site_data = all_site_daily_data[site_name]
                                match = site_data[site_data['date'] == date]
                                
                                if not match.empty:
                                    row_data.append(match.iloc[0]['Failed Count'])
                                else:
                                    row_data.append(0)
                            
                            ws.append(row_data)
                        
                        ws.column_dimensions['A'].width = 12
                        for col in range(2, len(headers) + 1):
                            ws.column_dimensions[chr(64 + col)].width = 15
                        
                        chart = LineChart()
                        chart.title = f"Daily Failure Count Trend - All Sites ({start_date_str} to {end_date_str})"
                        chart.style = 12
                        chart.y_axis.title = 'Number of Failed Equipment'
                        chart.x_axis.title = 'Date'
                        chart.x_axis.number_format = 'yyyy-mm-dd'
                        
                        num_dates = len(all_dates)
                        cats = Reference(ws, min_col=1, min_row=2, max_row=num_dates + 1)
                        
                        distinct_colors = [
                            "FF0000", "0000FF", "00AA00", "FF8C00", "9400D3", 
                            "00CED1", "DC143C", "FFD700"
                        ]
                        
                        for idx, site_name in enumerate(sorted(all_site_daily_data.keys())):
                            col_idx = 2 + idx
                            data = Reference(ws, min_col=col_idx, min_row=1, max_row=num_dates + 1)
                            chart.add_data(data, titles_from_data=True)
                            
                            if idx < len(distinct_colors):
                                chart.series[idx].graphicalProperties.line.solidFill = distinct_colors[idx]
                                chart.series[idx].graphicalProperties.line.width = 25000
                                chart.series[idx].marker.symbol = "circle"
                                chart.series[idx].marker.size = 5
                        
                        chart.set_categories(cats)
                        chart.height = 15
                        chart.width = 30
                        chart.legend.position = 'r'
                        
                        chart_position = f"A{num_dates + 4}"
                        ws.add_chart(chart, chart_position)
                
                log_message(f"Fail count report exported to: {export_path}", "INFO")
                
                num_sites = len(all_site_reports)
                site_list = ", ".join(sorted(all_site_reports.keys()))
                summary_msg = (
                    f"Successfully exported multi-site report to:\n{export_path}\n\n"
                    f"Sheets created:\n"
                    f"  • {num_sites} Site Report(s): {site_list}\n"
                    f"  • Consolidated Summary (daily failure trend chart)\n"
                    f"  • Date range: {start_date_str} to {end_date_str}"
                )
                
                self.after(0, lambda: messagebox.showinfo("Export Success", summary_msg))
        
        except Exception as export_err:
            log_message(f"Export error: {export_err}", "ERROR")
            self.after(0, lambda err=export_err: 
                messagebox.showerror("Export Error", str(err)))

    def show_test_counts_by_interval(self):
        """Show test counts by 30-minute intervals"""
        try:
            start_date_str = self.from_date.get_date().strftime('%Y-%m-%d')
            end_date_str = self.to_date.get_date().strftime('%Y-%m-%d')

            with self.get_db_connection() as conn:
                query = """
                SELECT equipment_id, time_tested
                FROM sensor_tests
                WHERE DATE(time_tested) BETWEEN DATE(?) AND DATE(?)
                """
                df = pd.read_sql(query, conn, params=[start_date_str, end_date_str])

            if df.empty:
                messagebox.showinfo("Interval View", 
                                  "No data to display in the selected date range.")
                return

            df['time_tested'] = pd.to_datetime(df['time_tested'])
            df['date'] = df['time_tested'].dt.date
            df['interval'] = df['time_tested'].dt.floor('30min')

            grouped = df.groupby(['equipment_id', 'date', 'interval']).size().reset_index(
                name='test_count'
            )

            interval_window = tk.Toplevel(self) 
            interval_window.title(
                f"Test Counts by 30-Minute Interval for {self.selected_shaft.get()}"
            )
            interval_window.geometry("800x600")

            scroll_frame = tk.Frame(interval_window)
            scroll_frame.pack(fill='both', expand=True)
            
            scrollbar = ttk.Scrollbar(scroll_frame)
            scrollbar.pack(side='right', fill='y')

            tree = ttk.Treeview(
                scroll_frame, 
                columns=('equipment_id', 'date', 'interval', 'test_count'), 
                show='headings', 
                yscrollcommand=scrollbar.set
            )
            scrollbar.config(command=tree.yview)
            
            style = ttk.Style()
            style.configure('Interval.Treeview.Heading', 
                            font=('Segoe UI', 10, 'bold'), 
                            background=self.colors['primary'], 
                            foreground='white')
            tree.config(style='Interval.Treeview')

            for col in tree['columns']:
                tree.heading(col, text=col.replace('_', ' ').title()) 
                tree.column(col, width=120)
            tree.pack(expand=True, fill='both')

            for _, row in grouped.iterrows():
                tree.insert('', 'end', values=(
                    row['equipment_id'], 
                    row['date'], 
                    row['interval'].strftime('%H:%M'), 
                    row['test_count']
                ))

        except ValueError as err:
            messagebox.showwarning("Error", str(err))
        except Exception as err:
            messagebox.showerror("Interval Display Error", str(err))


# ----------------------------------------------------------------------
# APPLICATION ENTRY POINT
# ----------------------------------------------------------------------
if __name__ == "__main__":
    try:
        log_message("="*50, "INFO")
        log_message("Starting Sentinel Audit Dashboard V1.4", "INFO")
        log_message("="*50, "INFO")
        app = SentinelDashboard()
        app.mainloop()
    except Exception as err:
        log_message(f"Critical application error: {err}", "ERROR")
        print(f"Application Error: {err}")