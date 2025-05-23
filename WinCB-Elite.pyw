# WinCB-Elite_part1.py
# Part 1 of 5 for WinCB-Elite.pyw
# Contains imports, configuration, and WinCB-Elite class initialization up to _setup_sidebar_buttons
# To combine: Concatenate part1 + part2 + part3 + part4 + part5 into WinCB-Elite.pyw

# -*- coding: utf-8 -*-
import customtkinter as ctk
from tkinter import Text, TclError, IntVar, Toplevel, Label, Frame, messagebox
import json
import os
import re
from PIL import Image, ImageDraw, ImageTk
import win32clipboard
import pystray
from pystray import MenuItem as item
import threading
import time
import traceback
import base64
import shutil
from io import BytesIO
import sys
import ctypes  # For DPI awareness
from pathlib import Path  # Added for better path handling

# --- Configuration ---
APP_NAME = "WinCB-Elite"
HISTORY_LIMIT = 50
AUTOSAVE_DELAY_MS = 1500  # Delay for auto-saving text edits (1.5 seconds)
MAIN_BG_COLOR = "#2b2b2b"  # Main dark background color
TEXT_FG_COLOR = "#ffffff"  # White text color
CURSOR_COLOR = "#ffffff"  # White cursor color
HIGHLIGHT_COLOR = "cyan"  # Highlight border for focused text widget

# --- Path Configuration ---
# Use user's home directory for persistent data
try:
    APP_DATA_DIR = Path.home() / "WinCB-Elite"
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    
    # Create dedicated directory for batch outputs
    BATCH_SAVE_DIR = APP_DATA_DIR / "batchoutputs"
    BATCH_SAVE_DIR.mkdir(parents=True, exist_ok=True)
    
    HISTORY_FILE_PATH = APP_DATA_DIR / "wincb-elite_history.json"
    CONFIG_FILE_PATH = APP_DATA_DIR / "wincb-elite_config.json"
    print(f"Using data directory: {APP_DATA_DIR}")
except Exception as path_e:
    print(f"FATAL: Could not create or access data directory: {Path.home() / 'WinCB-Elite'}")
    print(f"Error: {path_e}")
    # Fallback to script directory if home directory fails (less ideal for EXEs)
    try:
        script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    except NameError:
        script_dir = Path(os.getcwd())
    APP_DATA_DIR = script_dir
    HISTORY_FILE_PATH = APP_DATA_DIR / "wincb-elite_history.json"
    CONFIG_FILE_PATH = APP_DATA_DIR / "wincb-elite_config.json"
    
    # Create batch outputs directory even in fallback mode
    BATCH_SAVE_DIR = APP_DATA_DIR / "batchoutputs"
    try:
        BATCH_SAVE_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        # Last resort fallback to main directory
        BATCH_SAVE_DIR = APP_DATA_DIR
        
    print(f"Warning: Falling back to script directory for data: {APP_DATA_DIR}")

# Preview settings for Title Modal
PREVIEW_MAX_TEXT_LINES = 10
PREVIEW_MAX_TEXT_CHARS_PER_LINE = 50
PREVIEW_MAX_IMAGE_WIDTH = 150
PREVIEW_MAX_IMAGE_HEIGHT = 100

# Clipboard format constant
CF_DIB = 8  # Device Independent Bitmap

# Try importing CTkToolTip with a fallback
try:
    from customtkinter import CTkToolTip

    HAS_CTK_TOOLTIP = True
except ImportError:
    print("Info: CTkToolTip not available. Falling back to basic Tkinter tooltips.")
    HAS_CTK_TOOLTIP = False


class WinCB_Elite:
    """A clipboard manager application with history, text/image support, and system tray integration."""

    def __init__(self):
        """Initialize the WinCB-Elite application window and components."""
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.root = ctk.CTk()
        self.root.title(APP_NAME)
        self.root.geometry("900x700+650+0")
        self.root.minsize(750, 550)
        
        # Set window icon - try custom icon from config first, then fallback to default
        try:
            # First check if we have a saved custom icon in config
            self._load_config()  # Make sure config is loaded first
            custom_icon_path = self.config.get("custom_icon_path")
            
            if custom_icon_path and os.path.exists(custom_icon_path):
                # Use previously selected custom icon
                print(f"Using previously selected icon: {custom_icon_path}")
                try:
                    self.root.iconbitmap(custom_icon_path)
                except Exception as custom_err:
                    print(f"Error loading custom icon: {custom_err}")
                    # If custom icon fails, fall back to default
                    self.root.iconbitmap("icon.ico")
            else:
                # No custom icon, use default
                self.root.iconbitmap("icon.ico")
        except Exception:
            # If all icon attempts fail, silently continue - the app is more important
            pass
        
        self.root.resizable(True, True)
        self.root.lift()
        self.root.attributes("-topmost", True)
        self.root.after(
            1000,
            lambda: (
                self.root.attributes("-topmost", False)
                if getattr(self, "running", False)
                else None
            ),
        )
        self.root.protocol("WM_DELETE_WINDOW", self._prompt_close)

        # --- State Variables ---
        self.history = []
        self.filtered_history_indices = []
        self.current_filtered_index = -1
        self.last_clip_text = None
        self.last_clip_img_data = None
        self.ignore_clip_until = 0
        self.running = True
        self.capture_paused = False
        self.preview_popup = None
        self.current_display_image = None
        self.tk_image_references = {}
        self.save_timer_id = None
        self.current_clip_modified = False
        self.icon = None
        self.icon_thread = None
        self.additional_clipboard = None  # Stores additional copied content
        self.in_progress_clip = {"text": "", "images": []}  # In-progress clip construction
        self.click_outside_handler_id = None # For managing context menu click-outside binding
        self.current_group_name = None  # Tracks the loaded clip group name
        self.auto_pause_seconds = 0  # 0 means disabled, otherwise seconds until auto-pause
        self.auto_pause_timer = None  # Timer ID for auto-pause functionality
        self.last_activity_time = time.time()  # Track when the last clipboard activity occurred
        
        # Tag system - Color definitions
        self.TAG_COLORS = {
            "red": "#ff4d4d",     # Brighter red
            "blue": "#3399ff",    # Vibrant blue
            "green": "#33cc33",   # Brighter green
            "yellow": "#ffcc00",  # Golden yellow
            "purple": "#cc66ff"   # Brighter purple
        }
        self.current_clip_tags = []  # Tags for the currently displayed clip
        self.tag_colors = {}  # Map of tag name -> color name
        self._load_tag_colors()  # Load saved tag associations
        
        # --- Main Layout ---
        self.master_frame = ctk.CTkFrame(self.root, fg_color="transparent")
        self.master_frame.pack(fill="both", expand=True)

        # Calculate minimum height for sidebar_frame to ensure all buttons are visible
        separator_height = (
            20  # 10px pady top + 2px height + 5px pady bottom + 3px for label
        )
        button_space = 35  # 30px height + 5px pady
        num_buttons = 10  # Older, Newer, Delete, Copy Sel, Copy Clip, Preview, Clear, Hide, Export, Save/Restore
        num_separators = 3  # Navigation, Management, File Management
        bottom_buttons = 2  # Pause/Resume and Close
        bottom_space = (
            80  # 35px per button + 5px pady top for Pause, 5px pady bottom for Close
        )
        sidebar_min_height = (
            (num_buttons * button_space)
            + (num_separators * separator_height)
            + bottom_space
            + 40
        )  # Increased buffer
        self.sidebar_frame = ctk.CTkFrame(
            self.master_frame,
            width=180,
            height=sidebar_min_height,
            fg_color=MAIN_BG_COLOR,
        )
        self.sidebar_frame.pack(side="right", fill="y", padx=(15, 5), pady=5)
        self.sidebar_frame.pack_propagate(False)

        self.content_frame = ctk.CTkFrame(self.master_frame, fg_color="transparent")
        self.content_frame.pack(side="left", fill="both", expand=True, padx=5, pady=5)

        # --- Top Controls Area ---
        self.top_nav = ctk.CTkFrame(self.content_frame, fg_color="transparent")
        self.top_nav.pack(fill="x", pady=(0, 5), padx=5)
        self.top_nav.columnconfigure(1, weight=1)
        self.top_nav.rowconfigure(1, pad=3)  # Add padding between rows

        # Search frame
        self.search_frame = ctk.CTkFrame(self.top_nav, fg_color="transparent")
        self.search_frame.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        ctk.CTkLabel(self.search_frame, text="Search:").pack(side="left", padx=(0, 2))
        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            self.search_frame,
            textvariable=self.search_var, 
            width=120
        )
        self.search_entry.pack(side="left")
        # Direct binding to immediately force a refresh when search is cleared
        def on_search_clear(event):
            # If backspace/delete results in empty field, force refresh immediately
            if self.search_entry.get() == "":
                # Force show ALL clips immediately
                self.filtered_history_indices = list(range(len(self.history)))
                # Clear any filtering state
                self._is_tag_filtering = False
                if hasattr(self, '_current_filter_tag'):
                    delattr(self, '_current_filter_tag')
                # Update display
                self.current_filtered_index = 0
                self._show_clip()
                # Show confirmation that filter is cleared
                self._show_popup("Cleared all filters - showing all clips")
        
        self.search_entry.bind("<BackSpace>", on_search_clear)
        self.search_entry.bind("<Delete>", on_search_clear)
        self.search_entry.bind("<KeyRelease>", self._on_search_change)
        self._add_tooltip(self.search_entry, "Search clip titles and text content")
        
        # Add a dedicated clear button
        def force_clear_search():
            # Clear the search field
            self.search_var.set("")
            # Force show ALL clips
            self.filtered_history_indices = list(range(len(self.history)))
            # Reset tag filtering
            self._is_tag_filtering = False
            if hasattr(self, '_current_filter_tag'):
                delattr(self, '_current_filter_tag')
            # Update display
            self.current_filtered_index = 0
            self._show_clip()
            # Show confirmation
            self._show_popup("Showing all clips")
            
        self.clear_search_btn = ctk.CTkButton(
            self.search_frame,
            text="⌫",  # Backspace symbol is more intuitive for clearing
            width=28,
            height=30,
            command=force_clear_search
        )
        self.clear_search_btn.pack(side="left", padx=(5, 0))
        self._add_tooltip(self.clear_search_btn, "Clear search and show all clips")

        # Buffer functionality still exists but status indicator removed from UI
        # (Still accessible via Ctrl+Shift+C to copy, Ctrl+Shift+V to paste)
        self.buffer_status_var = ctk.StringVar(value="Buffer: Empty")  # Keep variable for internal use

        self.title_frame = ctk.CTkFrame(self.top_nav, fg_color="transparent")
        self.title_frame.grid(row=0, column=1, padx=5, pady=5, sticky="ew")
        ctk.CTkLabel(self.title_frame, text="Title:").pack(side="left", padx=(0, 2))
        self.title_var = ctk.StringVar()
        self.title_entry = ctk.CTkEntry(
            self.title_frame,
            textvariable=self.title_var,
            placeholder_text="Clip Title",
            state="disabled",
        )
        self.title_entry.pack(side="left", fill="x", expand=True, padx=(0, 5))
        self.title_entry.bind("<KeyRelease>", self._update_clip_title)
        self._add_tooltip(self.title_entry, "Edit the title (saved automatically)")

        self.page_label_placeholder = ctk.CTkFrame(
            self.title_frame, fg_color="transparent", width=100
        )
        self.page_label_placeholder.pack(side="right", padx=5)
        self.page_label = None
        
        # Add Group Name display row
        self.group_frame = ctk.CTkFrame(self.top_nav, fg_color="transparent", height=30)
        self.group_frame.grid(row=1, column=0, columnspan=3, sticky="ew", padx=5, pady=(0, 5))
        
        # Group label with variable to update
        self.group_name_var = ctk.StringVar(value="")
        self.group_label = ctk.CTkLabel(
            self.group_frame,
            textvariable=self.group_name_var,
            font=("Segoe UI", 13, "bold"),
            text_color="#4a95eb",  # Use a bright blue color that stands out
            anchor="w"
        )
        self.group_label.pack(side="left", fill="x", expand=True)
        self._update_group_display()  # Initialize the group display
        
        # Add Tags Frame
        self.tags_frame = ctk.CTkFrame(self.top_nav, fg_color="transparent", height=35)
        self.tags_frame.grid(row=2, column=0, columnspan=3, sticky="ew", padx=5, pady=(0, 5))
        self.tags_frame.grid_propagate(False)  # Prevent frame from expanding beyond its specified size
        
        # Label for tags section
        self.tags_label = ctk.CTkLabel(
            self.tags_frame,
            text="Tags:",
            font=("Segoe UI", 11),
            width=40,
            anchor="w"
        )
        self.tags_label.pack(side="left", padx=(0, 5))
        
        # Frame to hold tag buttons
        self.tag_buttons_frame = ctk.CTkFrame(self.tags_frame, fg_color="transparent", height=30)
        self.tag_buttons_frame.pack(side="left", fill="x", expand=True)
        self.tag_buttons_frame.pack_propagate(False)  # Prevent frame from expanding
        
        # Buttons frame for tag management
        tag_control_frame = ctk.CTkFrame(self.tags_frame, fg_color="transparent")
        tag_control_frame.pack(side="right")
        
        # Clear tags button (first/leftmost position)
        self.clear_tags_btn = ctk.CTkButton(
            tag_control_frame,
            text="🗑",  # Trash can icon better indicates removal
            width=22,
            height=30,
            font=("Segoe UI", 10),
            fg_color="#C0392B",  # Red color
            hover_color="#E74C3C",  # Lighter red on hover
            command=self._clear_current_clip_tags
        )
        self.clear_tags_btn.pack(side="left", padx=(0, 5))
        self._add_tooltip(self.clear_tags_btn, "Clear all tags from current clip")
        
        # Add tag button (second position)
        self.add_tag_btn = ctk.CTkButton(
            tag_control_frame,
            text="+",
            width=20,
            height=30,
            font=("Segoe UI", 15, "bold"),  # Larger font for + symbol
            fg_color="#1E5631",  # Dark green color
            hover_color="#2E7D32",  # Slightly lighter green for hover
            command=self._show_tag_dialog
        )
        self.add_tag_btn.pack(side="left", padx=(0, 5))  # Add right padding
        self._add_tooltip(self.add_tag_btn, "Add a new tag")
        
        # Refresh button (third/rightmost position) - reset all filters and show all clips
        self.refresh_btn = ctk.CTkButton(
            tag_control_frame,
            text="↻ Refresh",  # Added refresh symbol for clarity
            width=80,  # Slightly wider for the new symbol
            height=30,
            font=("Segoe UI", 10, "bold"),
            fg_color="#1E8449",  # Green color
            hover_color="#27AE60",  # Lighter green on hover
            command=self._reset_filtering
        )
        self.refresh_btn.pack(side="left")
        self._add_tooltip(self.refresh_btn, "Refresh view - show all clips")

        # --- Scrollable Content Area ---
        self.scrollable = ctk.CTkScrollableFrame(
            self.content_frame, fg_color=MAIN_BG_COLOR
        )
        self.scrollable.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        self.scrollable.bind("<Configure>", self._on_scrollable_configure)

        # Text Widget (Editable Area)
        self.textbox = Text(
            self.scrollable,
            wrap="word",
            state="disabled",
            borderwidth=0,
            highlightthickness=1,
            highlightbackground=MAIN_BG_COLOR,
            highlightcolor=HIGHLIGHT_COLOR,
            bg=MAIN_BG_COLOR,
            fg=TEXT_FG_COLOR,
            insertbackground=CURSOR_COLOR,
            font=("Segoe UI", 12),
            undo=True,
            selectbackground=HIGHLIGHT_COLOR,
            selectforeground="black",
        )
        self.textbox.bind("<KeyRelease>", self._on_text_edited)
        self.textbox.bind(
            "<FocusIn>", lambda e: self.textbox.config(highlightthickness=1)
        )
        self.textbox.bind(
            "<FocusOut>", lambda e: self.textbox.config(highlightthickness=0)
        )

        # Image Label
        self.img_label = Label(self.scrollable, bg=MAIN_BG_COLOR, anchor="nw")

        # Empty Message Label
        self.empty_message_label = ctk.CTkLabel(
            self.scrollable, text="", anchor="nw", justify="left", text_color="#aaaaaa"
        )

        # --- Sidebar Buttons ---
        self._setup_sidebar_buttons()

        # --- Initialization ---
        # Config was already loaded for icon, but we'll ensure everything is loaded
        if not hasattr(self, 'config'):
            self._load_config()  # Load config before history
        self._load_history()
        self._filter_and_show()
        self.root.after(500, self.poll_clipboard)
        
        # Start auto-pause timer if enabled
        if self.auto_pause_seconds > 0 and not self.capture_paused:
            self.root.after(1000, self._start_auto_pause_timer)
            print(f"Auto-pause timer scheduled: {self.auto_pause_seconds} seconds")
            
        self.icon_thread = threading.Thread(target=self._setup_tray, daemon=True)
        self.icon_thread.start()

        # Ensure context menu is initialized and bound correctly
        self.context_menu = ctk.CTkFrame(self.root, fg_color="gray25")
        self.context_menu.copy_btn = ctk.CTkButton(
            self.context_menu,
            text="Copy Active Clip",
            command=self.copy_active_clip_to_buffer,
            width=220 
        )
        self.context_menu.paste_btn = ctk.CTkButton(
            self.context_menu,
            text="Paste from Clipboard",
            command=self.paste_from_buffer_to_in_progress_clip,
            width=220
        )
        self.context_menu.clear_btn = ctk.CTkButton(
            self.context_menu,
            text="Clear Clipboard",
            command=self.clear_additional_buffer,
            width=220
        )
        self.context_menu.copy_btn.pack(pady=2)
        self.context_menu.paste_btn.pack(pady=2)
        self.context_menu.clear_btn.pack(pady=2)
        self.context_menu.place_forget()

        # Bind right-click events to show context menu
        self.textbox.bind("<Button-3>", self.show_context_menu)
        self.img_label.bind("<Button-3>", self.show_context_menu)

        self.root.bind("<Control-Shift-c>", lambda e: self.copy_focused_content_to_buffer())
        self.root.bind("<Control-Shift-C>", lambda e: self.copy_focused_content_to_buffer())
        self.root.bind("<Control-Shift-v>", lambda e: self.paste_from_buffer_to_in_progress_clip())
        self.root.bind("<Control-Shift-V>", lambda e: self.paste_from_buffer_to_in_progress_clip())
        
        # New bindings for in-progress clip
        self.root.bind("<Control-n>", lambda e: self.start_new_clip_from_selection())
        self.root.bind("<Control-a>", lambda e: self.add_selection_to_in_progress_clip())
        self.root.bind("<Control-A>", lambda e: self.add_selection_to_in_progress_clip())
        self.root.bind("<Control-s>", lambda e: self.save_in_progress_clip())
        self.root.bind("<Control-S>", lambda e: self.save_in_progress_clip())

        # Context-aware copy/paste bindings
        self.root.bind("<Control-c>", lambda e: self._context_aware_copy(e))
        self.root.bind("<Control-C>", lambda e: self._context_aware_copy(e))
        self.root.bind("<Control-v>", lambda e: self._context_aware_paste(e))
        self.root.bind("<Control-V>", lambda e: self._context_aware_paste(e))

        print("WinCB-Elite UI Initialized.")
        self.root.mainloop()

    # --- Setup Helpers ---

    def _setup_sidebar_buttons(self):
        """Creates and packs all buttons in the sidebar frame."""
        # Define button width and colors at the beginning to use throughout
        bottom_button_width = 160  # Width for all sidebar buttons
        navigation_button_color = "#3a6ea5"  # Blue-gray for navigation buttons
        management_button_color = "#555555"  # Mid-gray for management buttons
        
        self._add_sidebar_separator("Navigation")
        # Frame to hold the Start and End buttons side by side
        nav_jump_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        nav_jump_frame.pack(pady=5, padx=10, anchor="n")

        # End button (jump to oldest clip, last index) - Now on the left
        self.end_btn = ctk.CTkButton(
            nav_jump_frame, 
            text="|<", 
            width=55, 
            height=30, 
            fg_color=navigation_button_color,
            command=self.jump_to_oldest
        )
        self.end_btn.pack(side="left", padx=(0, 5))
        self._add_tooltip(self.end_btn, "Jump to oldest clip (Alt+End)")

        # Start button (jump to newest clip, index 0) - Now on the right
        self.start_btn = ctk.CTkButton(
            nav_jump_frame, 
            text=">|", 
            width=55, 
            height=30, 
            fg_color=navigation_button_color,
            command=self.jump_to_newest
        )
        self.start_btn.pack(side="right")
        self._add_tooltip(self.start_btn, "Jump to newest clip (Alt+Home)")

        # Older and Newer buttons
        # Navigation buttons use the color defined at the top
        
        self.older_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="\u2190 Older",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=navigation_button_color,
            command=self.prev_clip,
        )
        self.older_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.older_btn, "View previous clip in history (Alt+Left)")
        self.newer_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Newer \u2192",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=navigation_button_color,
            command=self.next_clip,
        )
        self.newer_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.newer_btn, "View next clip in history (Alt+Right)")
        self.delete_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Delete Clip",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=navigation_button_color,
            command=self.delete_current_clip,
        )
        self.delete_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.delete_btn, "Delete the currently displayed clip (Delete)"
        )

        self._add_sidebar_separator("Management")
        # Management buttons use the color defined at the top
        
        self.copy_sel_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Copy Selection",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=management_button_color,
            command=self.copy_selection_to_history,
        )
        self.copy_sel_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.copy_sel_btn, "Copy selected text from the editor as a new clip"
        )
        self.copy_clip_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Copy Clip",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=management_button_color,
            command=self.copy_clip_to_clipboard,
        )
        self.copy_clip_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.copy_clip_btn,
            "Copy the entire current clip back to the system clipboard",
        )
        self.titles_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Preview Titles",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=management_button_color,
            command=self._show_titles_modal,
        )
        self.titles_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.titles_btn, "View/select clips by title (with previews)")
        self.clear_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Clear History",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=management_button_color,
            command=self.clear_history,
        )
        self.clear_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.clear_btn, "Delete all clips from history (confirmation required)"
        )
        self.hide_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Hide",
            width=bottom_button_width,  # Match File Management button width
            height=30,
            fg_color=management_button_color,
            command=self._hide_window,
        )
        self.hide_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.hide_btn, "Minimize WinCB-Elite to the system tray")

        self._add_sidebar_separator("File Management")
        # File Management section uses the same width as defined at the top
        self.save_batch_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Readable Batch Export",
            width=bottom_button_width,
            height=30,
            command=self._prompt_and_save_batch,
        )
        self.save_batch_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.save_batch_btn,
            "Export currently filtered clips to a readable text file\n(For reference while working in other applications)",
        )
        self.save_restore_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Save/Restore Clip Group",
            width=bottom_button_width,
            height=30,
            command=self._prompt_save_restore_group,
        )
        self.save_restore_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.save_restore_btn,
            "Save or load clip groups as JSON files for project switching",
        )
        
        # Add button to change app icon
        self.change_icon_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Change App Icon",
            width=bottom_button_width,
            height=30,
            command=self._change_app_icon,
        )
        self.change_icon_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.change_icon_btn,
            "Change the application icon for system tray",
        )

        # Optional: Add Save In-Progress Clip button (can be hidden for simpler interface)
        if False:  # Set to True if you want this advanced feature
            self.save_in_progress_btn = ctk.CTkButton(
                self.sidebar_frame,
                text="Save In-Progress Clip",
                width=bottom_button_width,
                height=30,
                command=self.save_in_progress_clip,
            )
            self.save_in_progress_btn.pack(pady=5, padx=10, anchor="n")
            self._add_tooltip(
                self.save_in_progress_btn,
                "Save the current in-progress clip to history",
            )

        # Buttons at the very bottom, matching the width, with Close below Pause Capture
        # Create a frame to hold the bottom buttons in a vertical stack
        bottom_frame = ctk.CTkFrame(self.sidebar_frame, fg_color=MAIN_BG_COLOR)
        bottom_frame.pack(side="bottom", padx=10, pady=(5, 5))

        self.pause_resume_btn = ctk.CTkButton(
            bottom_frame,
            text="Pause Capture",
            width=bottom_button_width,
            height=30,
            command=self._toggle_capture,
        )
        self.pause_resume_btn.pack(pady=(0, 5))
        self._add_tooltip(self.pause_resume_btn, "Pause/resume clipboard monitoring")

        self.close_btn = ctk.CTkButton(
            bottom_frame,
            text="Close",
            width=bottom_button_width,
            height=30,
            command=self._prompt_close,
            fg_color="#8B0000",
            hover_color="#A00000",
        )
        self.close_btn.pack()
        self._add_tooltip(self.close_btn, "Close WinCB-Elite (confirmation required)")

        # --- Key Bindings ---
        self.root.bind("<Alt-Left>", lambda e: self.prev_clip())
        self.root.bind("<Alt-Right>", lambda e: self.next_clip())
        # Add bindings for Start (newest) and End (oldest)
        self.root.bind("<Alt-Home>", lambda e: self.jump_to_newest())
        self.root.bind("<Alt-End>", lambda e: self.jump_to_oldest())

        # Conditional Delete key binding to avoid deleting clips while editing
        def delete_clip_if_no_focus(event):
            focused = self.root.focus_get()
            
            # NEVER EVER delete clips if search field has ANY involvement whatsoever
            if focused == self.search_entry:
                print("DEBUG: Delete key blocked - search entry has direct focus")
                return "break"  # Completely stop event propagation
            
            # Double-check search entry has no focus before continuing
            has_search_focus = False
            try:
                if self.search_entry.focus_get() or self.search_entry.focus_displayof():
                    has_search_focus = True
            except:
                pass
                
            if has_search_focus:
                return "break"
                
            # Only proceed if not editing text and search doesn't have focus
            if focused not in (self.title_entry, self.textbox, self.search_entry):
                self.delete_current_clip()

        self.root.bind("<Delete>", delete_clip_if_no_focus)
        self.root.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.root.bind("<Control-F>", lambda e: self.search_entry.focus_set())
        
        # Escape key to clear filtering
        self.root.bind("<Escape>", lambda e: self._reset_filtering())
        # Also bind Escape to the search entry for immediate clearing
        self.search_entry.bind("<Escape>", lambda e: self._reset_filtering())
        
        # Override the entry's handling of Delete/Backspace to prevent any focus issues
        self.search_entry.bind("<FocusIn>", lambda e: print("DEBUG: Search entry got focus"))
        self.search_entry.bind("<FocusOut>", lambda e: print("DEBUG: Search entry lost focus"))

    def _on_search_change(self, event=None):
        """Handle search input changes."""
        self._finalize_text_edit()
        
        # Update activity time when searching
        self._update_activity_time()
        
        # CRITICAL SAFETY CHECK
        # Detect Delete key and block it from triggering anything destructive
        if event and event.keysym == "Delete":
            print("DEBUG: Delete key pressed in search field - BLOCKED from processing")
            return  # Block any processing of the Delete key in search
            
        # Detect Backspace key deletion events
        is_backspace = event and event.keysym == "BackSpace"
        
        # Get current search text
        current_search = self.search_entry.get().strip()
        
        # Print more diagnostic info about the search change
        print(f"DEBUG: Search field changed: '{current_search}' via {event.keysym if event else 'unknown'}")
        
        # Store the current clip's original index before filtering changes
        current_original_index = -1
        if 0 <= self.current_filtered_index < len(self.filtered_history_indices):
            try:
                current_original_index = self.filtered_history_indices[self.current_filtered_index]
                # Store clip information for debugging
                clip_info = "unknown"
                if 0 <= current_original_index < len(self.history):
                    clip = self.history[current_original_index]
                    clip_info = f"{clip.get('title', '?')} ({clip.get('type', '?')})"
                print(f"DEBUG: Current clip before search change: {clip_info} at index {current_original_index}")
            except Exception as e:
                print(f"ERROR storing current clip info: {e}")
                pass
        
        # Check if we're in tag filtering mode
        was_tag_filtering = False
        current_tag_name = None
        if hasattr(self, '_is_tag_filtering') and self._is_tag_filtering:
            was_tag_filtering = True
            current_tag_name = getattr(self, '_current_filter_tag', None)
            print(f"DEBUG: In tag filtering mode for tag: {current_tag_name}")
            
            # When clearing the search box while in tag filtering mode
            if not current_search:
                print("DEBUG: Tag filter being cleared - showing all clips without changing anything")
                
                # Clear the tag filtering flags
                self._is_tag_filtering = False
                if hasattr(self, '_current_filter_tag'):
                    delattr(self, '_current_filter_tag')
                    
                # Reset to show all clips
                self.filtered_history_indices = list(range(len(self.history)))
                print(f"DEBUG: Reset to show all {len(self.filtered_history_indices)} clips")
                
                # Keep the same clip selected if possible
                if current_original_index >= 0:
                    if current_original_index < len(self.history):
                        try:
                            # The indices now match the original history, so the index remains the same
                            self.current_filtered_index = current_original_index
                            print(f"DEBUG: Kept original clip at index {current_original_index}")
                        except Exception as e:
                            print(f"ERROR finding original clip: {e}")
                            self.current_filtered_index = 0
                    else:
                        self.current_filtered_index = 0
                
                # Show the clip and return early to avoid normal filtering
                self._show_clip()
                self._show_popup("Cleared tag filter - showing all clips")
                return
            
        # Clear tag filtering on empty search - even if we weren't previously tag filtering
        elif not current_search:
            # Make sure any tag filtering is cleared
            self._is_tag_filtering = False
            if hasattr(self, '_current_filter_tag'):
                delattr(self, '_current_filter_tag')
                print("DEBUG: Cleared tag filter state on empty search")
                
        # Still using the tag filter - just with a modified search term
        elif current_search.startswith('#') and current_tag_name:
            if current_search[1:].lower() == current_tag_name.lower():
                # The tag search hasn't changed, just keep it as is
                pass
            else:
                # Changed to a different tag - update the filter tag
                self._current_filter_tag = current_search[1:]
                print(f"DEBUG: Changed tag filter to: {self._current_filter_tag}")
        else:
            # Changed to a normal search - no longer filtering by tag
            self._is_tag_filtering = False
            if hasattr(self, '_current_filter_tag'):
                delattr(self, '_current_filter_tag')
                print("DEBUG: Exited tag filtering mode - switched to normal search")
        
        # If we're not handling a special tag case, continue with normal filtering
        self._filter_and_show()

    # WinCB-Elite_part2.py
    # Part 2 of 5 for WinCB-Elite.pyw
    # Contains sidebar helpers and history persistence methods
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into WinCB-Elite.pyw

    def _add_sidebar_separator(self, label_text):
        """Add a separator with a label to the sidebar."""
        sep = ctk.CTkFrame(self.sidebar_frame, height=2, fg_color="#555555")
        sep.pack(fill="x", pady=(10, 2), padx=10, anchor="n")
        lbl = ctk.CTkLabel(
            self.sidebar_frame,
            text=label_text,
            font=("Segoe UI", 11, "bold"),
            text_color="#AAAAAA",
        )
        lbl.pack(pady=(0, 5), anchor="n")

    def _add_tooltip(self, widget, message):
        """Add a tooltip to a widget, using CTkToolTip or fallback to Tkinter."""
        if HAS_CTK_TOOLTIP:
            try:
                CTkToolTip(widget, message=message, delay=0.5)
            except Exception:
                self._add_tkinter_tooltip(widget, message)
        else:
            self._add_tkinter_tooltip(widget, message)

    def _add_tkinter_tooltip(self, widget, message):
        """Fallback basic Tkinter tooltip."""
        tooltip, show_timer = None, None

        def schedule(e):
            nonlocal show_timer
            if show_timer:
                try:
                    widget.after_cancel(show_timer)
                except:
                    pass
            if widget.winfo_exists():
                show_timer = widget.after(500, lambda ev=e: show(ev))

        def show(e):
            nonlocal tooltip
            if (
                not widget.winfo_exists()
                or tooltip
                or not getattr(self, "running", True)
            ):
                return
            try:
                x, y = (
                    widget.winfo_rootx() + 20,
                    widget.winfo_rooty() + widget.winfo_height() + 5,
                )
            except TclError:
                return
            tooltip = Toplevel(widget)
            tooltip.wm_overrideredirect(True)
            tooltip.wm_geometry(f"+{x}+{y}")
            tooltip.attributes("-topmost", True)
            Label(
                tooltip,
                text=message,
                justify="left",
                bg="#333",
                fg="#fff",
                relief="solid",
                borderwidth=1,
                font=("Segoe UI", 9),
                padx=5,
                pady=3,
            ).pack(ipadx=1)

        def hide(e=None):
            nonlocal tooltip, show_timer
            if show_timer:
                try:
                    widget.after_cancel(show_timer)
                except:
                    pass
                show_timer = None
            if tooltip:
                try:
                    tooltip.destroy()
                except:
                    pass
                tooltip = None

        widget.bind("<Enter>", schedule, add="+")
        widget.bind("<Leave>", hide, add="+")
        widget.bind("<Button-1>", hide, add="+")
        widget.bind("<Unmap>", hide, add="+")
        widget.bind("<Destroy>", hide, add="+")

    # --- History Persistence ---

    def _load_history(self):
        """Load clipboard history from JSON file in the APP_DATA_DIR."""
        if not HISTORY_FILE_PATH.exists():
            print(f"History file not found: {HISTORY_FILE_PATH}")
            return
        try:
            with open(HISTORY_FILE_PATH, "r", encoding="utf-8") as f:
                saved_data = json.load(f)

            loaded_count = 0
            temp_history = []
            if not isinstance(saved_data, list):
                print("Warning: History file does not contain a list. Creating backup.")
                self._backup_corrupted_history()
                return

            for item_data in saved_data:
                if not isinstance(item_data, dict) or not all(
                    k in item_data for k in ("type", "content", "title")
                ):
                    print(f"Skipping invalid history item: {item_data}")
                    continue

                entry = {"type": item_data["type"], "title": item_data.get("title", "")}
                content = item_data["content"]
                timestamp = item_data.get("timestamp", 0)
                tags = item_data.get("tags", [])  # Load tags from the saved data

                if entry["type"] == "image":
                    try:
                        if isinstance(content, str):
                            entry["content"] = base64.b64decode(content)
                        else:
                            print(
                                f"Skipping image entry with non-string content: '{entry['title']}'"
                            )
                            continue
                    except (base64.binascii.Error, TypeError, ValueError) as e:
                        print(
                            f"Skipping invalid base64 image data '{entry['title']}': {e}"
                        )
                        continue
                elif entry["type"] == "text":
                    if isinstance(content, str):
                        entry["content"] = content
                    else:
                        print(f"Skipping non-string text entry: '{entry['title']}'")
                        continue
                else:
                    print(f"Skipping entry with unknown type: {entry['type']}")
                    continue

                entry["timestamp"] = timestamp
                entry["tags"] = tags  # Add tags to the entry
                temp_history.append(entry)
                loaded_count += 1

            temp_history.sort(key=lambda x: x.get("timestamp", 0), reverse=True)
            self.history = temp_history[:HISTORY_LIMIT]
            print(f"Loaded {len(self.history)} items from {HISTORY_FILE_PATH}.")
            if loaded_count > HISTORY_LIMIT:
                print(f"History truncated to the latest {HISTORY_LIMIT} items.")

        except json.JSONDecodeError as e:
            print(f"Error decoding history file: {HISTORY_FILE_PATH} - {e}")
            self._backup_corrupted_history()
        except Exception as e:
            print(f"Error loading history from {HISTORY_FILE_PATH}: {e}")
            traceback.print_exc()
            self._backup_corrupted_history()
            
    def _load_config(self):
        """Load application configuration from JSON file."""
        self.config = {
            "custom_icon_path": None,  # Default: no custom icon
            "auto_pause_seconds": 0    # Default: auto-pause disabled
        }
        
        if not CONFIG_FILE_PATH.exists():
            print(f"Config file not found, using defaults: {CONFIG_FILE_PATH}")
            return
            
        try:
            with open(CONFIG_FILE_PATH, "r", encoding="utf-8") as f:
                saved_config = json.load(f)
                
            if isinstance(saved_config, dict):
                # Update config with saved values
                self.config.update(saved_config)
                print(f"Loaded configuration from {CONFIG_FILE_PATH}")
                
                # Validate custom icon path if one is set
                if self.config.get("custom_icon_path"):
                    icon_path = self.config["custom_icon_path"]
                    if not os.path.exists(icon_path):
                        print(f"Warning: Custom icon not found at {icon_path}")
                        self.config["custom_icon_path"] = None
                        
                # Load auto-pause setting if it exists
                if "auto_pause_seconds" in self.config:
                    try:
                        seconds = int(self.config["auto_pause_seconds"])
                        if seconds >= 0:
                            self.auto_pause_seconds = seconds
                            print(f"Loaded auto-pause setting: {seconds} seconds")
                    except (ValueError, TypeError) as e:
                        print(f"Warning: Invalid auto-pause setting: {e}")
                        self.config["auto_pause_seconds"] = 0
            else:
                print("Warning: Config file does not contain a dictionary.")
                
        except json.JSONDecodeError as e:
            print(f"Error decoding config file: {CONFIG_FILE_PATH} - {e}")
        except Exception as e:
            print(f"Error loading config from {CONFIG_FILE_PATH}: {e}")
            traceback.print_exc()
            
    def _save_config(self):
        """Save application configuration to JSON file."""
        if not self.running:
            return
            
        try:
            temp_file_path = CONFIG_FILE_PATH.with_suffix(".tmp")
            with open(temp_file_path, "w", encoding="utf-8") as f:
                json.dump(self.config, f, ensure_ascii=False, indent=2)
                
            os.replace(temp_file_path, CONFIG_FILE_PATH)
            print(f"Saved configuration to {CONFIG_FILE_PATH}")
            
        except Exception as e:
            print(f"Error saving config to {CONFIG_FILE_PATH}: {e}")
            traceback.print_exc()
            try:
                if temp_file_path.exists():
                    temp_file_path.unlink()
            except Exception as te:
                print(f"Warning: Could not remove temporary config file {temp_file_path}: {te}")

    def _backup_corrupted_history(self):
        """Backup corrupted history file to the APP_DATA_DIR."""
        if HISTORY_FILE_PATH.exists():
            try:
                backup_path = (
                    APP_DATA_DIR / f"{HISTORY_FILE_PATH.name}.backup_{int(time.time())}"
                )
                shutil.copy(str(HISTORY_FILE_PATH), str(backup_path))
                print(f"Corrupted history file backed up to: {backup_path}")
            except Exception as e:
                print(f"Error backing up corrupted history file: {e}")

    def _save_history(self):
        """Save clipboard history to JSON file in the APP_DATA_DIR."""
        if not self.running:
            return
        try:
            history_to_save = sorted(
                self.history, key=lambda x: x.get("timestamp", 0), reverse=True
            )
            saveable_history = []

            for item in history_to_save[:HISTORY_LIMIT]:
                save_item = {
                    "type": item.get("type"),
                    "title": item.get("title", ""),
                    "timestamp": item.get("timestamp", 0),
                    "tags": item.get("tags", []),  # Add tags to saved data
                }
                content = item.get("content")

                if save_item["type"] == "image" and isinstance(content, bytes):
                    save_item["content"] = base64.b64encode(content).decode("utf-8")
                elif save_item["type"] == "text" and isinstance(content, str):
                    save_item["content"] = content
                else:
                    print(
                        f"Warning: Skipping invalid item during save: Type={save_item['type']}, Title='{save_item['title']}'"
                    )
                    continue

                if save_item["type"] and save_item["content"] is not None:
                    saveable_history.append(save_item)

            temp_file_path = HISTORY_FILE_PATH.with_suffix(".tmp")
            with open(temp_file_path, "w", encoding="utf-8") as f:
                json.dump(saveable_history, f, ensure_ascii=False, indent=2)

            os.replace(temp_file_path, HISTORY_FILE_PATH)

        except Exception as e:
            print(f"Error saving history to {HISTORY_FILE_PATH}: {e}")
            traceback.print_exc()
            try:
                if temp_file_path.exists():
                    temp_file_path.unlink()
            except Exception as te:
                print(
                    f"Warning: Could not remove temporary save file {temp_file_path}: {te}"
                )

    # WinCB-Elite_part3.py
    # Part 3 of 5 for WinCB-Elite.pyw
    # Contains clipboard polling and filtering methods
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into WinCB-Elite.pyw

    # --- Clipboard Polling & History Management ---

    def _toggle_capture(self):
        """Toggle clipboard capture on/off."""
        self.capture_paused = not self.capture_paused
        default_color = ctk.ThemeManager.theme["CTkButton"]["fg_color"]
        hover_color = ctk.ThemeManager.theme["CTkButton"]["hover_color"]

        if self.capture_paused:
            self.pause_resume_btn.configure(
                text="Resume Capture", fg_color="orange", hover_color="darkorange"
            )
            print("Clipboard capture PAUSED. Use 'Resume Capture' button or tray menu to resume.")
            
            # Cancel auto-pause timer if active
            if self.auto_pause_timer:
                self.root.after_cancel(self.auto_pause_timer)
                self.auto_pause_timer = None
        else:
            self.pause_resume_btn.configure(
                text="Pause Capture", fg_color=default_color, hover_color=hover_color
            )
            print("Clipboard capture RESUMED. Now monitoring clipboard changes.")
            
            # Start auto-pause timer if enabled
            if self.auto_pause_seconds > 0:
                self._start_auto_pause_timer()

    def poll_clipboard(self):
        """Poll system clipboard for new content with improved error handling."""
        if not self.running:
            return

        now = time.time()

        if self.capture_paused:
            if self.running:
                self.root.after(1000, self.poll_clipboard)
            return

        # More robust ignore period check
        if now < self.ignore_clip_until:
            remaining = self.ignore_clip_until - now
            if remaining > 0.5:
                print(f"DEBUG: Ignoring clipboard for {remaining:.2f}s more")
            if self.running:
                self.root.after(200, self.poll_clipboard)
            return

        new_clip_found = False
        clip_type = None
        clip_data = None

        try:
            # Use a more robust approach with retries
            max_retries = 3
            retry_count = 0
            
            while retry_count < max_retries:
                try:
                    win32clipboard.OpenClipboard()
                    break
                except Exception as e:
                    retry_count += 1
                    if retry_count >= max_retries:
                        print(f"DEBUG: Failed to open clipboard after {max_retries} attempts: {e}")
                        if self.running:
                            self.root.after(1000, self.poll_clipboard)
                        return
                    time.sleep(0.1)  # Short delay before retry
            
            has_unicode = win32clipboard.IsClipboardFormatAvailable(
                win32clipboard.CF_UNICODETEXT
            )
            has_text = win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_TEXT)

            if has_unicode or has_text:
                text_data = ""
                if has_unicode:
                    try:
                        text_data = win32clipboard.GetClipboardData(
                            win32clipboard.CF_UNICODETEXT
                        )
                    except TypeError as te:
                        print(f"DEBUG: CF_UNICODETEXT failed ({te}), trying CF_TEXT.")
                        if has_text:
                            try:
                                text_data = win32clipboard.GetClipboardData(
                                    win32clipboard.CF_TEXT
                                ).decode("mbcs")
                            except Exception as decode_err:
                                print(f"DEBUG: CF_TEXT decode failed: {decode_err}")
                        else:
                            text_data = None
                    except Exception as e:
                        print(f"DEBUG: Error getting CF_UNICODETEXT: {e}")
                        text_data = None
                elif has_text:
                    try:
                        text_data = win32clipboard.GetClipboardData(
                            win32clipboard.CF_TEXT
                        ).decode("mbcs")
                    except Exception as e:
                        print(f"DEBUG: Error getting/decoding CF_TEXT: {e}")

                if (
                    text_data is not None
                    and text_data.strip()
                    and text_data != self.last_clip_text
                ):
                    # Avoid excessively long text clips
                    if len(text_data) > 500_000:
                        print(
                            f"DEBUG: Skipped excessively long text clip ({len(text_data)} chars)"
                        )
                    else:
                        print(f"DEBUG: New text detected ({len(text_data)} chars)")
                        self.last_clip_text = text_data
                        self.last_clip_img_data = None
                        new_clip_found = True
                        clip_type = "text"
                        clip_data = text_data

            if not new_clip_found and win32clipboard.IsClipboardFormatAvailable(CF_DIB):
                try:
                    image_data = win32clipboard.GetClipboardData(CF_DIB)
                    if image_data:  # Check if we got valid data first
                        if image_data != self.last_clip_img_data:
                            # Only print detected when it's actually NEW
                            print(f"DEBUG: New image data detected: {len(image_data)} bytes")
                            self.last_clip_img_data = image_data
                            self.last_clip_text = None  # Important to clear the other type
                            new_clip_found = True
                            clip_type = "image"
                            clip_data = image_data
                except Exception as e:
                    print(f"DEBUG: Error getting CF_DIB data: {e}")
                    traceback.print_exc()

            win32clipboard.CloseClipboard()

            if new_clip_found and clip_type and clip_data is not None:
                log_preview = ""
                if clip_type == "text":
                    log_preview = clip_data[:60].replace("\n", "\\n").replace(
                        "\r", ""
                    ) + ("..." if len(clip_data) > 60 else "")
                    print(f"DEBUG: Captured text: {log_preview}")
                elif clip_type == "image":
                    print(f"DEBUG: Captured image ({len(clip_data)} bytes)")
                action = lambda ct=clip_type, cd=clip_data: self._add_to_history(ct, cd)
                self.root.after(0, action)

        except win32clipboard.error as e:
            if e.winerror != 5:
                print(f"DEBUG: Clipboard access error (winerror {e.winerror}): {e}")
            try:
                win32clipboard.CloseClipboard()
            except:
                pass
        except Exception as e:
            print(f"DEBUG: Error during clipboard polling: {e}")
            traceback.print_exc()
            try:
                win32clipboard.CloseClipboard()
            except:
                pass
        finally:
            if self.running:
                self.root.after(1000, self.poll_clipboard)

    def _add_to_history(self, data_type, data, is_from_selection=False):
        """Adds new clipboard content or selection to history."""
        print(f"DEBUG: _add_to_history called with type={data_type}, data_type={type(data)}, is_from_selection={is_from_selection}")
        
        if not self.running or data is None:
            print(
                f"DEBUG: Failed to add to history: running={self.running}, data={data is None}"
            )
            return

        # Update last activity time whenever content is added to history
        self.last_activity_time = time.time()
        
        current_time = time.time()
        new_entry = {
            "type": data_type,
            "content": data,
            "title": "",
            "timestamp": current_time,
            "tags": [],  # Initialize empty tags list for new clips
        }

        is_duplicate = False
        if not is_from_selection and self.history:
            last_entry = self.history[0]
            last_ts = last_entry.get("timestamp", 0)
            if current_time - last_ts < 0.8:
                is_duplicate = True
                print(
                    f"DEBUG: Skipped: Too soon after last capture ({current_time - last_ts:.2f}s)"
                )
            elif data_type == last_entry.get("type"):
                if (
                    data_type == "text"
                    and isinstance(data, str)
                    and isinstance(last_entry.get("content"), str)
                ):
                    if data.strip() == last_entry["content"].strip():
                        is_duplicate = True
                        print("DEBUG: Skipped: Duplicate text content")
                elif (
                    data_type == "image"
                    and isinstance(data, bytes)
                    and data == last_entry.get("content")
                ):
                    is_duplicate = True
                    print("DEBUG: Skipped: Duplicate image content")

        if is_duplicate:
            return

        if data_type == "image" and isinstance(data, bytes):
            try:
                img = Image.open(BytesIO(data))
                size_str = f"{img.width}x{img.height}"
                timestamp_str = time.strftime("%Y%m%d_%H%M%S")
                new_entry["title"] = (
                    f"Image_{timestamp_str}_{size_str}"  # More descriptive prefix
                )
                print(f"DEBUG: Image title generated: {new_entry['title']}")
            except Exception as e:
                print(f"DEBUG: Failed to process image for title: {e}")
                traceback.print_exc()
                new_entry["title"] = f"Image_{time.strftime('%Y%m%d_%H%M%S')}_[ERR]"
                
        elif data_type == "text" and isinstance(data, str):
            stripped_data = data.strip()
            if not stripped_data:
                print("DEBUG: Skipped empty text entry.")
                return
            first_line = stripped_data.split("\n", 1)[0]
            max_title_len = 60  # Increased to capture more context
            title = re.sub(r"\s+", " ", first_line).strip()
            new_entry["title"] = (
                (title[:max_title_len] + "...") if len(title) > max_title_len else title
            )
            # Remove trailing periods or ellipses for clarity
            new_entry["title"] = new_entry["title"].rstrip(".").rstrip("...")
            print(f"DEBUG: Text title generated: {new_entry['title']}")
            
        elif data_type == "mixed" and isinstance(data, dict):
            # Handle mixed type (dict with text and image keys)
            print(f"DEBUG: Processing mixed type with keys: {list(data.keys())}")
            
            # Validate mixed content structure
            if not ('text' in data or 'image' in data):
                print(f"DEBUG: Invalid mixed content structure: {data.keys()}")
                return
                
            text_content = data.get('text', '')
            if text_content:
                if not isinstance(text_content, str):
                    print(f"DEBUG: Mixed clip has invalid text type: {type(text_content)}")
                    text_content = str(text_content)
                    
                # Generate title from text content
                first_line = text_content.strip().split("\n", 1)[0]
                max_title_len = 50
                title = re.sub(r"\s+", " ", first_line).strip()
                new_entry["title"] = (
                    (title[:max_title_len] + "...") if len(title) > max_title_len else title
                )
                # Add image indicator if it has an image
                if 'image' in data and data['image']:
                    new_entry["title"] = f"[Mixed] {new_entry['title']}"
                new_entry["title"] = new_entry["title"].rstrip(".").rstrip("...")
            else:
                # No text, just use timestamp with mixed indicator
                timestamp_str = time.strftime("%Y%m%d_%H%M%S")
                new_entry["title"] = f"Mixed_{timestamp_str}"
                
            print(f"DEBUG: Mixed title generated: {new_entry['title']}")
            
        else:
            print(
                f"DEBUG-ERROR: Invalid data type '{data_type}' or data provided to _add_to_history."
            )
            return

        print(f"DEBUG: Successfully created new entry with title: {new_entry['title']}")
        
        if len(self.history) >= HISTORY_LIMIT:
            print(f"DEBUG: History limit reached ({HISTORY_LIMIT}), removing oldest entry")
            self.history.pop()

        self.history.insert(0, new_entry)
        print(f"Added '{new_entry['title']}' to history (Total: {len(self.history)})")
        self._save_history()

        search_term = self.search_entry.get().lower().strip()
        should_update_display_fully = True

        if search_term:
            matches_search = False
            try:
                if search_term in new_entry["title"].lower():
                    matches_search = True
                    print(f"DEBUG: New entry title matches search: '{search_term}'")
                elif (
                    new_entry["type"] == "text"
                    and search_term in new_entry["content"].lower()
                ):
                    matches_search = True
                    print(f"DEBUG: New entry text content matches search: '{search_term}'")
            except Exception as search_check_err:
                print(
                    f"DEBUG-ERROR: Error checking if new item matches search: {search_check_err}"
                )

            if not matches_search:
                should_update_display_fully = False
                self._filter_history()
                self._update_page_label()
                print(
                    "DEBUG: New item added but doesn't match current filter. Display not changed."
                )

        if should_update_display_fully:
            self._filter_history()
            if self.filtered_history_indices and self.filtered_history_indices[0] == 0:
                self.current_filtered_index = 0
                self._show_clip()
            else:
                self._filter_and_show()

    # --- Filtering and Display ---

    def _filter_history(self):
        """Filter history based on search query."""
        search_term = self.search_entry.get().lower().strip()
        
        # Handle empty search - SHOW ALL CLIPS
        if not search_term:
            print("REFRESH: Showing all clips - search field empty")
            
            # Remember current clip if possible
            current_clip_index = -1
            if 0 <= self.current_filtered_index < len(self.filtered_history_indices):
                try:
                    current_clip_index = self.filtered_history_indices[self.current_filtered_index]
                except:
                    pass
            
            # ALWAYS show all clips when search is empty
            self.filtered_history_indices = list(range(len(self.history)))
            
            # Reset tag filtering state
            self._is_tag_filtering = False
            if hasattr(self, '_current_filter_tag'):
                delattr(self, '_current_filter_tag')
            
            # Try to keep same clip selected
            if 0 <= current_clip_index < len(self.history):
                self.current_filtered_index = current_clip_index
            else:
                self.current_filtered_index = 0 if self.filtered_history_indices else -1
                
            return
        
        print(f"DEBUG: Filtering history with search term: '{search_term}'")
        self.filtered_history_indices = []
        
        # If it's a tag search with # prefix, look specifically in tags
        if search_term.startswith('#') and len(search_term) > 1:
            tag_to_find = search_term[1:].lower()
            print(f"DEBUG: Searching for specific tag: '{tag_to_find}'")
            for i, item in enumerate(self.history):
                if "tags" in item:
                    for tag in item.get("tags", []):
                        if tag.lower() == tag_to_find:
                            self.filtered_history_indices.append(i)
                            break
        # Otherwise do a normal search across all fields
        else:
            for i, item in enumerate(self.history):
                item_matches = False
                try:
                    # Check title
                    if search_term in item.get("title", "").lower():
                        item_matches = True
                    # Check content for text clips
                    elif item.get("type") == "text" and isinstance(item.get("content"), str):
                        if search_term in item["content"].lower():
                            item_matches = True
                    # Check tags
                    if not item_matches and "tags" in item:
                        for tag in item.get("tags", []):
                            if search_term in tag.lower():
                                item_matches = True
                                break
                    
                    if item_matches:
                        self.filtered_history_indices.append(i)
                except Exception as e:
                    print(f"WARNING: Error during search filter on item index {i}: {e}")

    def _filter_and_show(self):
        """Filter history and update display, trying to preserve current item if possible."""
        # Store original clip info before filtering
        current_original_index = -1
        current_clip_info = None
        
        if 0 <= self.current_filtered_index < len(self.filtered_history_indices):
            try:
                current_original_index = self.filtered_history_indices[
                    self.current_filtered_index
                ]
                
                # Store more information about the current clip for debug purposes
                if 0 <= current_original_index < len(self.history):
                    clip = self.history[current_original_index]
                    current_clip_info = {
                        "title": clip.get("title", ""),
                        "type": clip.get("type", ""),
                        "tags": clip.get("tags", []),
                        "index": current_original_index
                    }
                    print(f"DEBUG: Current clip before filtering: {current_clip_info}")
            except IndexError:
                pass

        # Apply filtering
        self._filter_history()
        
        # Check if we have any clips after filtering
        num_filtered = len(self.filtered_history_indices)
        print(f"DEBUG: After filtering, found {num_filtered} matching clips")

        if num_filtered == 0:
            # No clips match the filter
            self.current_filtered_index = -1
            print("DEBUG: No clips match the current filter")
        else:
            # Try to find the same clip in the new filtered list
            new_filtered_index = -1
            if current_original_index != -1:
                try:
                    # Find the same clip in the new filtered list
                    new_filtered_index = self.filtered_history_indices.index(
                        current_original_index
                    )
                    print(f"DEBUG: Found current clip at new index {new_filtered_index} in filtered list")
                except ValueError:
                    print(f"DEBUG: Current clip (index {current_original_index}) not found in new filtered list")
                    if current_clip_info:
                        print(f"DEBUG: Missing clip details: {current_clip_info}")
                    pass
                    
            # Update the current index (keep the same clip if possible, otherwise show first result)
            self.current_filtered_index = (
                new_filtered_index if new_filtered_index != -1 else 0
            )
        
        # Show the clip
        self._show_clip()
        
    def _reset_filtering(self):
        """Reset all filtering to show all clips - use as an emergency escape from stuck filters."""
        print("RESET: Emergency reset of all filtering")
        
        # Clear tag filtering state
        self._is_tag_filtering = False
        if hasattr(self, '_current_filter_tag'):
            delattr(self, '_current_filter_tag')
            
        # Clear search box WITHOUT triggering filter changes
        self.search_var.set("")
        self.search_entry.delete(0, "end")
        
        # FORCE show all clips
        self.filtered_history_indices = list(range(len(self.history)))
        
        # ALWAYS go to the first clip (index 0) when refreshing
        self.current_filtered_index = 0 if self.filtered_history_indices else -1
            
        # Show the clip and confirmation
        self._show_clip()
        self._show_popup("Cleared all filters - showing first clip")

    def _update_page_label(self):
        """Update the page label like 'Clip 3/15'."""
        num_filtered = len(self.filtered_history_indices)
        label_text = "No Clips"

        if num_filtered > 0:
            if self.current_filtered_index >= 0:
                label_text = f"Clip {self.current_filtered_index + 1}/{num_filtered}"
            else:
                label_text = f"? / {num_filtered}"
        elif self.search_entry.get().strip():
            label_text = "0 Matches"
        elif not self.history:
            label_text = "Empty"

        try:
            if self.page_label is None or not self.page_label.winfo_exists():
                for widget in self.page_label_placeholder.winfo_children():
                    widget.destroy()
                self.page_label = ctk.CTkLabel(
                    self.page_label_placeholder, text=label_text, anchor="e"
                )
                self.page_label.pack(side="right", fill="x", expand=True)
            else:
                self.page_label.configure(text=label_text)
        except Exception as e:
            print(f"Warning: Failed to update page label: {e}")

    # WinCB-Elite_part4.py
    # Part 4 of 5 for WinCB-Elite.pyw
    # Contains display, text editing, button actions, and titles modal methods
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into WinCB-Elite.pyw

    def _create_main_thumb(self, image_bytes):
        """Creates a PhotoImage thumbnail from bytes for the main display area."""
        if not image_bytes:
            return None
        try:
            img = Image.open(BytesIO(image_bytes))
            thumb = img.copy()
            # Use fixed minimum dimensions to prevent shrinking
            max_width = max(self.scrollable.winfo_width() - 20, 500)
            max_height = max(self.scrollable.winfo_height() - 20, 400)
            thumb.thumbnail((max_width, max_height), Image.Resampling.LANCZOS)
            photo = ImageTk.PhotoImage(thumb)
            return photo
        except Exception as e:
            print(f"Error creating main display thumbnail: {e}")
            traceback.print_exc()
            return None

    def _show_clip(self):
        """Displays the currently selected clip in the content area."""
        if not self.running:
            return
            
        # Update activity time when displaying a clip
        self._update_activity_time()

        self.current_clip_modified = False
        self.current_display_image = None

        if self.save_timer_id:
            self.root.after_cancel(self.save_timer_id)
            self.save_timer_id = None

        try:
            if self.textbox.winfo_ismapped():
                self.textbox.pack_forget()
            if self.img_label.winfo_ismapped():
                self.img_label.pack_forget()
            if self.empty_message_label.winfo_ismapped():
                self.empty_message_label.pack_forget()

            self.img_label.config(image="")
            self.img_label.image = None

            self.textbox.configure(state="normal")
            self.textbox.delete("1.0", "end")
            self.textbox.configure(state="disabled")
            self.textbox.edit_reset()
            self.textbox.edit_modified(False)
            self.textbox.tag_remove("search_highlight", "1.0", "end")

        except Exception as e:
            print(f"Warning: Error resetting display area: {e}")

        display_message = None
        can_edit_title = False
        can_edit_content = False
        item_to_display = None

        search_active = self.search_entry.get().strip()

        if not self.history:
            display_message = "Clipboard history is empty."
        elif not self.filtered_history_indices:
            display_message = (
                f"No clips found matching '{search_active}'."
                if search_active
                else "No clips available."
            )
        elif not (
            0 <= self.current_filtered_index < len(self.filtered_history_indices)
        ):
            display_message = "Error: Invalid display index."
            print(
                f"ERROR: Invalid current_filtered_index: {self.current_filtered_index}, "
                f"filtered_history_indices length: {len(self.filtered_history_indices)}"
            )
            self.current_filtered_index = -1
        else:
            try:
                original_index = self.filtered_history_indices[
                    self.current_filtered_index
                ]
                if 0 <= original_index < len(self.history):
                    item_to_display = self.history[original_index]
                else:
                    display_message = "Error: History data inconsistency."
                    print(
                        f"ERROR: Original index {original_index} out of bounds for history length {len(self.history)}"
                    )
                    self._filter_history()
                    self.current_filtered_index = (
                        0 if self.filtered_history_indices else -1
                    )

            except IndexError:
                display_message = "Error: Could not retrieve clip."
                print(
                    f"ERROR: IndexError accessing filtered_history_indices at {self.current_filtered_index}"
                )
                self.current_filtered_index = -1
            except Exception as e:
                display_message = f"Error loading clip details: {e}"
                print(f"Error retrieving clip item: {e}")
                traceback.print_exc()

        if item_to_display:
            clip_type = item_to_display.get("type")
            clip_content = item_to_display.get("content")
            clip_title = item_to_display.get("title", "Untitled")

            self.title_var.set(clip_title)
            can_edit_title = True

            if clip_type == "text" and isinstance(clip_content, str):
                try:
                    self.textbox.configure(state="normal")
                    self.textbox.insert("1.0", clip_content)
                    self.textbox.configure(state="normal")
                    self.textbox.edit_reset()
                    self.textbox.edit_modified(False)
                    can_edit_content = True

                    if search_active:
                        start_index = "1.0"
                        while True:
                            match_pos = self.textbox.search(
                                search_active,
                                start_index,
                                stopindex="end",
                                nocase=1,
                                count=IntVar(),
                            )
                            if not match_pos:
                                break
                            match_len = len(search_active)
                            end_index = f"{match_pos}+{match_len}c"
                            if match_len == 0:
                                break
                            try:
                                self.textbox.tag_add(
                                    "search_highlight", match_pos, end_index
                                )
                            except TclError as tag_err:
                                print(f"Warning: Error adding highlight tag: {tag_err}")
                                break
                            start_index = end_index
                        self.textbox.tag_config(
                            "search_highlight", background="orange", foreground="black"
                        )

                    self.textbox.pack(fill="both", expand=True, pady=5, padx=5)

                except Exception as text_e:
                    print(f"Error displaying text content: {text_e}")
                    display_message = "Error displaying text."
                    if self.textbox.winfo_ismapped():
                        self.textbox.pack_forget()
                    can_edit_content = False

            elif clip_type == "image" and isinstance(clip_content, bytes):
                self.current_display_image = self._create_main_thumb(clip_content)
                if self.current_display_image:
                    try:
                        self.img_label.config(image=self.current_display_image)
                        self.img_label.image = self.current_display_image
                        self.img_label.pack(anchor="nw", pady=5, padx=5)
                        
                        # Always show textbox for image clips to make it editable
                        self.textbox.configure(state="normal")
                        self.textbox.delete("1.0", "end")
                        self.textbox.pack(fill="both", expand=True, pady=5, padx=5)
                        can_edit_content = True
                    except Exception as img_e:
                        print(f"Error setting image label: {img_e}")
                        display_message = "Error displaying image preview."
                        self.current_display_image = None
                        if self.img_label.winfo_ismapped():
                            self.img_label.pack_forget()
                else:
                    display_message = "Image preview unavailable or invalid."
                    self.textbox.configure(state="disabled")

            elif clip_type == "mixed" and isinstance(clip_content, dict):
                # Handle mixed clip type (containing both text and image)
                text_content = clip_content.get("text", "")
                image_data = clip_content.get("image")
                
                # Display image if available
                if image_data and isinstance(image_data, bytes):
                    self.current_display_image = self._create_main_thumb(image_data)
                    if self.current_display_image:
                        try:
                            self.img_label.config(image=self.current_display_image)
                            self.img_label.image = self.current_display_image
                            self.img_label.pack(anchor="nw", pady=5, padx=5)
                        except Exception as img_e:
                            print(f"Error setting image in mixed clip: {img_e}")
                
                # Display and enable text editing
                try:
                    self.textbox.configure(state="normal")
                    self.textbox.delete("1.0", "end")
                    if text_content:
                        self.textbox.insert("1.0", text_content)
                    self.textbox.pack(fill="both", expand=True, pady=5, padx=5)
                    can_edit_content = True
                except Exception as text_e:
                    print(f"Error displaying text in mixed clip: {text_e}")

            else:
                display_message = f"Unsupported clip type: {clip_type}"
                self.textbox.configure(state="disabled")

        if display_message:
            self.empty_message_label.configure(text=display_message)
            self.empty_message_label.pack(fill="x", anchor="nw", pady=20, padx=20)
            self.title_var.set("")
            can_edit_title = False
            self.textbox.configure(state="disabled")

        self.title_entry.configure(state="normal" if can_edit_title else "disabled")
        self._update_page_label()
        self._update_tag_display()  # Update tag display with current clip's tags
        self.root.update_idletasks()
        self._update_scrollregion()

        # Make text editable regardless of clip type
        if self.textbox.winfo_ismapped() and 0 <= self.current_filtered_index < len(self.filtered_history_indices):
            self.textbox.configure(state="normal")

    # --- Text Editing ---

    def _on_text_edited(self, event=None):
        """Handles text modification in the textbox, schedules auto-save."""
        # Update activity time when actively editing
        self._update_activity_time()
        
        if (
            not self.running
            or self.textbox.cget("state") == "disabled"
            or not self.textbox.edit_modified()
        ):
            return

        self.current_clip_modified = True

        if self.save_timer_id:
            self.root.after_cancel(self.save_timer_id)

        self.save_timer_id = self.root.after(AUTOSAVE_DELAY_MS, self._save_edited_text)

    def _save_edited_text(self):
        """Saves edited text from the textbox to the corresponding history item."""
        self.save_timer_id = None

        if (
            not self.running
            or not self.current_clip_modified
            or self.textbox.cget("state") == "disabled"
        ):
            return

        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            print("Error saving text: Invalid current index.")
            return

        try:
            original_index = self.filtered_history_indices[self.current_filtered_index]
            if not (0 <= original_index < len(self.history)):
                print("Error saving text: Original history index out of bounds.")
                return

            new_content = self.textbox.get("1.0", "end-1c")
            clip_type = self.history[original_index]["type"]
            
            if clip_type == "text":
                # Standard text clip update
                if self.history[original_index]["content"] != new_content:
                    print(f"Auto-saving changes to clip: {self.history[original_index]['title']}")
                    self.history[original_index]["content"] = new_content
                    self.history[original_index]["timestamp"] = time.time()
                    self._save_history()
                    self.textbox.edit_modified(False)
                    self.current_clip_modified = False
                else:
                    self.textbox.edit_modified(False)
                    self.current_clip_modified = False
            
            elif clip_type == "image":
                # For image clips, convert to mixed type if text is added
                if new_content.strip():
                    print(f"Converting image clip to mixed type with text: {self.history[original_index]['title']}")
                    image_content = self.history[original_index]["content"]
                    # Create new mixed content
                    mixed_content = {"text": new_content, "image": image_content}
                    # Update the clip to mixed type
                    self.history[original_index]["type"] = "mixed"
                    self.history[original_index]["content"] = mixed_content
                    self.history[original_index]["timestamp"] = time.time()
                    self._save_history()
                
                self.textbox.edit_modified(False)
                self.current_clip_modified = False
                
            elif clip_type == "mixed":
                # Update text portion of mixed clip
                current_content = self.history[original_index]["content"]
                if current_content.get("text") != new_content:
                    print(f"Auto-saving changes to mixed clip: {self.history[original_index]['title']}")
                    current_content["text"] = new_content
                    self.history[original_index]["timestamp"] = time.time()
                    self._save_history()
                
                self.textbox.edit_modified(False)
                self.current_clip_modified = False
            else:
                print(f"Warning: Attempted to save text edit to unsupported item type: {clip_type}")
                self.textbox.edit_modified(False)
                self.current_clip_modified = False

        except IndexError:
            print("Error saving text: Index out of bounds (race condition?).")
        except Exception as e:
            print(f"Error saving edited text: {e}")
            traceback.print_exc()

    def _finalize_text_edit(self):
        """Immediately saves any pending text edits if the timer is active."""
        # Update activity time if there were text edits
        if self.current_clip_modified:
            self._update_activity_time()
            
        if self.save_timer_id:
            self.root.after_cancel(self.save_timer_id)
            self.save_timer_id = None
            if self.current_clip_modified:
                self._save_edited_text()
        elif self.current_clip_modified:
            self._save_edited_text()

    # --- Title Editing ---

    def _update_clip_title(self, event=None):
        """Updates the title of the current clip immediately on KeyRelease."""
        # Update activity time when editing titles
        self._update_activity_time()
        
        if not self.running or not (
            0 <= self.current_filtered_index < len(self.filtered_history_indices)
        ):
            return

        try:
            original_index = self.filtered_history_indices[self.current_filtered_index]
            if not (0 <= original_index < len(self.history)):
                print("Error updating title: Original history index out of bounds.")
                return

            new_title = self.title_var.get()

            if self.history[original_index]["title"] != new_title:
                self.history[original_index]["title"] = new_title
                self.history[original_index]["timestamp"] = time.time()
                self._save_history()

        except IndexError:
            print("Warning: Index error during title update (likely filter change).")
        except Exception as e:
            print(f"Error updating clip title: {e}")
            traceback.print_exc()

    # --- Button Actions ---

    def jump_to_oldest(self):
        """Moves to the oldest clip in the filtered history."""
        self._finalize_text_edit()
        self._update_activity_time()  # Update activity when navigating
        num_filtered = len(self.filtered_history_indices)
        if num_filtered <= 1:
            self._show_popup("No older clips available.")
            return
        self.current_filtered_index = num_filtered - 1  # Last index is the oldest
        self._show_clip()

    def jump_to_newest(self):
        """Moves to the newest clip in the filtered history."""
        self._finalize_text_edit()
        self._update_activity_time()  # Update activity when navigating
        num_filtered = len(self.filtered_history_indices)
        if num_filtered <= 1:
            self._show_popup("No newer clips available.")
            return
        self.current_filtered_index = 0  # Index 0 is the newest
        self._show_clip()

    def delete_current_clip(self):
        """Deletes the currently displayed clip from history."""
        self._finalize_text_edit()
        self._update_activity_time()  # Update activity when deleting clips

        if not self.running or not (
            0 <= self.current_filtered_index < len(self.filtered_history_indices)
        ):
            self._show_popup("No clip selected to delete.")
            return

        try:
            original_index_to_delete = self.filtered_history_indices[
                self.current_filtered_index
            ]

            if not (0 <= original_index_to_delete < len(self.history)):
                self._show_popup("Error deleting clip: Data inconsistency.")
                return
                
            # Get clip title for the confirmation message
            clip_title = self.history[original_index_to_delete].get("title", "Untitled clip")
            
            # Create confirmation dialog
            confirm_dialog = ctk.CTkToplevel(self.root)
            confirm_dialog.title("Confirm Delete")
            confirm_dialog.geometry("400x180")
            confirm_dialog.transient(self.root)
            confirm_dialog.grab_set()
            confirm_dialog.attributes("-topmost", True)
            confirm_dialog.configure(fg_color=MAIN_BG_COLOR)
            
            message = f"Are you sure you want to delete this clip?\n\n\"{clip_title}\""
            ctk.CTkLabel(
                confirm_dialog, 
                text=message,
                wraplength=350,
                justify="center"
            ).pack(pady=(20, 15), padx=20)
            
            buttons_frame = ctk.CTkFrame(confirm_dialog, fg_color="transparent")
            buttons_frame.pack(pady=(0, 15))
            
            def on_cancel():
                confirm_dialog.destroy()
                
            def on_confirm():
                confirm_dialog.destroy()
                
                # Now actually delete the clip
                try:
                    # Double-check the index is still valid
                    if 0 <= original_index_to_delete < len(self.history):
                        removed_item = self.history.pop(original_index_to_delete)
                        print(f"Deleted clip: {removed_item.get('title', '?')}")
                        self._save_history()
                        self._filter_and_show()
                    else:
                        self._show_popup("Error: The clip to delete no longer exists.")
                except Exception as delete_error:
                    print(f"Error during confirmed deletion: {delete_error}")
                    traceback.print_exc()
            
            ctk.CTkButton(
                buttons_frame,
                text="Cancel",
                width=100,
                command=on_cancel
            ).pack(side="left", padx=10)
            
            ctk.CTkButton(
                buttons_frame,
                text="Delete",
                width=100,
                fg_color="#8B0000",  # Dark red
                hover_color="#A00000",  # Slightly lighter red
                command=on_confirm
            ).pack(side="left", padx=10)
            
            # Center the dialog
            self._center_toplevel(confirm_dialog)

        except IndexError:
            self._show_popup("Error deleting clip: Invalid index.")
        except Exception as e:
            print(f"Error deleting clip: {e}")
            traceback.print_exc()

    def copy_selection_to_history(self):
        """Copies selected text or the current image as a new history item."""
        # Update activity time when creating new clips
        self._update_activity_time()
        # First check if we have text selected in the textbox
        if self.textbox.winfo_ismapped() and self.textbox.cget("state") == "normal":
            try:
                has_selection = self.textbox.tag_ranges("sel")
                if has_selection:
                    selection = self.textbox.get("sel.first", "sel.last")
                    if selection and selection.strip():
                        # Reset clipboard internal state to ensure fresh copy
                        self.last_clip_text = None
                        self.last_clip_img_data = None
                        
                        self._add_to_history("text", selection, is_from_selection=True)
                        self._force_copy_to_clipboard("text", selection)
                        
                        # Navigate to the new clip
                        self._filter_history()
                        self.current_filtered_index = 0
                        self._show_clip()
                        
                        # Ensure textbox remains editable
                        self.textbox.configure(state="normal")
                        self._show_popup("Selected text added as new clip and copied to clipboard.")
                        return
            except TclError:
                pass  # No selection
                
        # If no text selection, check for an image
        if self.img_label.winfo_ismapped() and self.current_display_image:
            # An image is displayed, let's copy it
            if (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
                try:
                    original_index = self.filtered_history_indices[self.current_filtered_index]
                    if (0 <= original_index < len(self.history) and 
                        self.history[original_index]["type"] == "image"):
                        # Create a new clip with this image
                        image_content = self.history[original_index]["content"]
                        self._add_to_history("image", image_content, is_from_selection=True)
                        self._force_copy_to_clipboard("image", image_content)
                        
                        # Navigate to the new clip
                        self._filter_history()
                        self.current_filtered_index = 0
                        self._show_clip()
                        
                        self._show_popup("Image copied to new clip and to clipboard.")
                        return
                except Exception as e:
                    print(f"Error copying image: {e}")
                    
        # If we get here, neither condition was met
        self._show_popup("No text selection or valid image found to copy.")
        # Make sure textbox is still editable if it's showing
        if self.textbox.winfo_ismapped():
            self.textbox.configure(state="normal")

    def copy_clip_to_clipboard(self):
        """Copies the entire current clip to the system clipboard."""
        if not self.running or not (
            0 <= self.current_filtered_index < len(self.filtered_history_indices)
        ):
            self._show_popup("No clip selected to copy.")
            return
        try:
            original_index = self.filtered_history_indices[self.current_filtered_index]
            if not (0 <= original_index < len(self.history)):
                self._show_popup("Error copying clip: Invalid index.")
                return

            item = self.history[original_index]
            clip_type = item.get("type")
            content = item.get("content")  # This line was missing

            if content is None:
                self._show_popup("Cannot copy: No content available.")
                return

            self._force_copy_to_clipboard(clip_type, content)
            self._show_popup(f"{clip_type.capitalize()} clip copied to clipboard.")

        except IndexError:
            self._show_popup("Error copying clip: Index out of bounds.")
        except Exception as e:
            print(f"Error copying clip to clipboard: {e}")
            traceback.print_exc()
            self._show_popup(f"Error copying clip: {e}")

    def _force_copy_to_clipboard(self, kind, content):
        """Helper to forcefully set the system clipboard with retry mechanism."""
        print(f"DEBUG: _force_copy_to_clipboard called with kind={kind}, content_size={len(content) if content else 0}")
        max_retries = 3
        retry_delay = 100  # ms
        
        for attempt in range(max_retries):
            try:
                # First try to get current clipboard state to check if we need to retry
                try:
                    win32clipboard.OpenClipboard()
                    win32clipboard.CloseClipboard()
                except Exception as e:
                    print(f"DEBUG: Clipboard busy on pre-check, will retry: {e}")
                    if attempt < max_retries - 1:
                        print(f"DEBUG: Waiting {retry_delay}ms before retry {attempt+1}/{max_retries}")
                        time.sleep(retry_delay/1000)
                        retry_delay *= 2  # Exponential backoff
                        continue
                    else:
                        raise  # Re-raise if this was the last attempt
                
                # Now do the actual clipboard operation
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                
                if kind == "text" and isinstance(content, str):
                    win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, content)
                    self.last_clip_text = content
                    self.last_clip_img_data = None
                    print(f"DEBUG: Text copied to clipboard successfully ({len(content)} chars)")
                elif kind == "image" and isinstance(content, bytes):
                    win32clipboard.SetClipboardData(CF_DIB, content)
                    self.last_clip_img_data = content
                    self.last_clip_text = None
                    print(f"DEBUG: Image copied to clipboard successfully ({len(content)} bytes)")
                elif kind == "mixed" and isinstance(content, dict):
                    # For mixed content, prioritize text for system clipboard
                    if "text" in content and content["text"]:
                        win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, content["text"])
                        self.last_clip_text = content["text"]
                        self.last_clip_img_data = None
                        print(f"DEBUG: Mixed (text) copied to clipboard successfully ({len(content['text'])} chars)")
                    elif "image" in content and content["image"]:
                        win32clipboard.SetClipboardData(CF_DIB, content["image"])
                        self.last_clip_img_data = content["image"]
                        self.last_clip_text = None
                        print(f"DEBUG: Mixed (image) copied to clipboard successfully ({len(content['image'])} bytes)")
                else:
                    print(f"DEBUG-ERROR: Cannot force copy to clipboard: Invalid type {kind} or content")
                    win32clipboard.CloseClipboard()
                    return False
                
                win32clipboard.CloseClipboard()
                
                # Only set ignore period for automatic clipboard monitoring
                # Use a shorter ignore period for manual operations to allow rapid paste sequences
                self.ignore_clip_until = time.time() + 0.5  # Reduced from 1.5s to 0.5s
                
                print(f"DEBUG: Clipboard operation completed successfully on attempt {attempt+1}")
                return True
                
            except Exception as e:
                print(f"DEBUG-ERROR: Error forcing clipboard copy (attempt {attempt+1}/{max_retries}): {e}")
                print(f"DEBUG-ERROR: Error type: {type(e).__name__}")
                traceback.print_exc()
                
                try:
                    win32clipboard.CloseClipboard()
                except:
                    pass
                    
                if attempt < max_retries - 1:
                    print(f"DEBUG: Waiting {retry_delay}ms before retry {attempt+1}/{max_retries}")
                    time.sleep(retry_delay/1000)
                    retry_delay *= 2  # Exponential backoff
                else:
                    self._show_popup(f"Clipboard Error: {e}", priority=2)
                    return False
        
        return False  # Should not reach here, but just in case

    def prev_clip(self):
        """Moves to the previous (older) clip."""
        self._finalize_text_edit()
        num_filtered = len(self.filtered_history_indices)
        if num_filtered <= 1:
            self._show_popup("No older clips available.")
            return
        if self.current_filtered_index < num_filtered - 1:
            self.current_filtered_index += 1
            self._show_clip()
        else:
            self._show_popup("Reached the oldest clip.")

    def next_clip(self):
        """Moves to the next (newer) clip."""
        self._finalize_text_edit()
        num_filtered = len(self.filtered_history_indices)
        if num_filtered <= 1:
            self._show_popup("No newer clips available.")
            return
        if self.current_filtered_index > 0:
            self.current_filtered_index -= 1
            self._show_clip()
        else:
            self._show_popup("Reached the newest clip.")

    def clear_history(self):
        """Confirms and clears the entire history."""
        self._finalize_text_edit()
        if not self.history:
            self._show_popup("Clipboard history is already empty.")
            return
        was_hidden = not self.root.winfo_viewable()
        if was_hidden:
            self._do_show_window()
        self.root.after(150, lambda wh=was_hidden: self._confirm_clear(wh))

    def _confirm_clear(self, was_hidden):
        """Shows the confirmation dialog for clearing history."""
        if not self.running:
            return
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Confirm Clear History")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)
        
        # Strong warning message
        ctk.CTkLabel(
            dialog,
            text="⚠️ WARNING ⚠️",
            font=("Segoe UI", 14, "bold"),
            text_color="#ff9900",
        ).pack(pady=(15, 5))
        
        ctk.CTkLabel(
            dialog,
            text="You are about to delete ALL clips in your clipboard history.\n\nThis action CANNOT be undone and will permanently remove all clips and tags.",
            wraplength=350,
            justify="center",
        ).pack(pady=(5, 20), padx=10)
        
        # Confirmation checkbox
        confirm_var = ctk.IntVar(value=0)
        confirm_check = ctk.CTkCheckBox(
            dialog,
            text="I understand I will lose all my clips and tags",
            variable=confirm_var,
            checkbox_width=20,
            checkbox_height=20,
            corner_radius=3
        )
        confirm_check.pack(pady=(0, 10))
        
        btns = ctk.CTkFrame(dialog, fg_color="transparent")
        btns.pack(pady=10)

        def on_yes():
            # First check if the confirmation checkbox is checked
            if confirm_var.get() != 1:
                # Show error if checkbox not checked
                error_label = ctk.CTkLabel(
                    dialog,
                    text="You must check the confirmation box",
                    text_color="#ff3333",
                    font=("Segoe UI", 11, "bold")
                )
                # Insert above button frame
                error_label.pack_forget()  # Ensure it's not already packed
                error_label.pack(before=btns, pady=(0, 5))
                # Highlight the checkbox
                confirm_check.configure(border_color="#ff3333", border_width=2)
                # Auto-close error after a few seconds
                dialog.after(3000, lambda: error_label.destroy())
                return
                
            try:
                print("Clearing clipboard history...")
                self.history = []
                self.filtered_history_indices = []
                self.current_filtered_index = -1
                self.tk_image_references.clear()
                self.search_var.set("")
                self.title_var.set("")
                # Clear group name when history is cleared
                self.current_group_name = None
                self._update_group_display()
                try:
                    win32clipboard.OpenClipboard()
                    win32clipboard.EmptyClipboard()
                    win32clipboard.CloseClipboard()
                    self.last_clip_text = None
                    self.last_clip_img_data = None
                    print("System clipboard cleared.")
                except Exception as clip_e:
                    print(f"Warning: Failed to clear system clipboard: {clip_e}")
                    try:
                        win32clipboard.CloseClipboard()
                    except:
                        pass
                self._save_history()
                self._show_clip()
                self._show_popup("Clipboard history cleared.")
            except Exception as e:
                print(f"Error clearing history: {e}")
                traceback.print_exc()
            finally:
                try:
                    dialog.destroy()
                except:
                    pass
                if was_hidden and self.running:
                    self.root.after(100, self._hide_window)

        def on_no():
            try:
                dialog.destroy()
            except:
                pass
            if was_hidden and self.running:
                self.root.after(100, self._hide_window)

        ctk.CTkButton(
            btns,
            text="Yes, Clear",
            width=110,
            command=on_yes,
            fg_color="#8B0000",
            hover_color="#A00000",
        ).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="No", width=110, command=on_no).pack(
            side="left", padx=10
        )
        self._center_toplevel(dialog)

    # --- Titles Modal ---

    def _show_titles_modal(self):
        """Opens a modal to preview/select clips."""
        self._finalize_text_edit()
        if not self.running or not self.history:
            self._show_popup("No clips available to preview.")
            return
        modal = ctk.CTkToplevel(self.root)
        modal.title("Preview Clip Titles")
        modal.geometry("350x450")
        modal.transient(self.root)
        modal.grab_set()
        modal.attributes("-topmost", True)
        modal.bind("<Destroy>", lambda e: self.tk_image_references.clear())
        modal.configure(fg_color=MAIN_BG_COLOR)

        main = ctk.CTkFrame(modal, fg_color=MAIN_BG_COLOR)
        main.pack(fill="both", expand=True, padx=5, pady=5)
        ctk.CTkLabel(
            main, text="Select clip (Hover for preview):", font=("Segoe UI", 12, "bold")
        ).pack(pady=(5, 10))
        scroll = ctk.CTkScrollableFrame(main, fg_color=MAIN_BG_COLOR)
        scroll.pack(fill="both", expand=True, padx=5, pady=(0, 5))
        for idx, item in enumerate(self.history):
            title = item.get("title", "").strip()
            clip_type = item.get("type")
            content = item.get("content")
            # Fallback title if empty or not descriptive
            if not title or len(title) < 5 or "ERR" in title:
                if clip_type == "text" and isinstance(content, str):
                    snippet = (
                        content.strip()
                        .split("\n", 1)[0][:30]
                        .replace("\n", " ")
                        .replace("\r", "")
                    )
                    title = (
                        f"Text: {snippet}" if snippet else f"Text Clip {idx+1} (Empty)"
                    )
                elif clip_type == "image":
                    title = f"Image Clip {idx+1}"
                else:
                    title = f"Clip {idx+1} (Unknown)"
            # Truncate long titles for display
            max_display_len = 50
            display_title = (
                (title[:max_display_len] + "...")
                if len(title) > max_display_len
                else title
            )
            btn = ctk.CTkButton(
                scroll,
                text=display_title,
                anchor="w",
                fg_color="#3a3a3a",
                hover_color="#4a4a4a",
                command=lambda i=idx, m=modal: self._select_clip_from_modal(i, m),
            )
            btn.pack(fill="x", pady=3, padx=5)
            
            # Use a delay mechanism to prevent multiple popups when moving between items
            def delayed_show(event, index=idx, widget=btn):
                # Cancel any previous scheduled preview
                if hasattr(self, 'preview_update_id') and self.preview_update_id:
                    try:
                        self.root.after_cancel(self.preview_update_id)
                    except:
                        pass
                
                # Schedule new preview after a short delay
                self.preview_update_id = self.root.after(
                    150,  # 150ms delay prevents popup flicker
                    lambda: self._show_preview_popup(event, index, widget)
                )
            
            # Improved hover handling
            btn.bind("<Enter>", delayed_show, add="+")
            btn.bind("<Leave>", lambda ev: self._hide_preview_popup(), add="+")
            
            # Ensure preview is hidden when clicking
            btn.bind("<Button-1>", lambda ev: self._hide_preview_popup(), add="+")
        self._center_toplevel(modal)

    def _show_preview_popup(self, event, history_index, widget):
        """Creates and displays a preview popup for clip content on hover."""
        if not self.running or not (0 <= history_index < len(self.history)):
            return
        
        # Track which widget triggered this preview to handle mouse movement
        self.last_preview_widget = widget
        
        # Cancel any pending preview update
        if hasattr(self, 'preview_update_id') and self.preview_update_id:
            try:
                self.root.after_cancel(self.preview_update_id)
                self.preview_update_id = None
            except Exception:
                pass
                
        # Ensure all existing preview popups are destroyed
        self._hide_preview_popup()
        
        # Force removal of any orphaned previews
        for child in self.root.winfo_children():
            try:
                if (isinstance(child, ctk.CTkToplevel) and 
                    child.winfo_exists() and 
                    child != self.root and 
                    hasattr(child, 'title') and 
                    child.title() == "Clip Preview"):
                    child.destroy()
            except Exception:
                pass
        
        try:
            # Get clip info
            clip = self.history[history_index]
            clip_type = clip.get("type", "unknown")
            content = clip.get("content")
            title = clip.get("title", "Untitled")
            
            if content is None:
                return
                
            # Create popup window
            self.preview_popup = ctk.CTkToplevel(self.root)
            self.preview_popup.title("Clip Preview")
            self.preview_popup.attributes("-topmost", True)
            self.preview_popup.overrideredirect(True)
            self.preview_popup.configure(fg_color=MAIN_BG_COLOR)
            
            # Add unique identifier to help with cleanup
            self.preview_popup.preview_id = time.time()
            
            # Add a callback to ensure this popup is tracked for cleanup
            self.preview_popup.bind("<Destroy>", lambda e: self._on_preview_destroy(e))
            
            # Position popup near mouse
            x, y = event.x_root, event.y_root
            
            # Create content frame with border for visibility
            content_frame = ctk.CTkFrame(
                self.preview_popup, 
                fg_color=MAIN_BG_COLOR,
                border_width=1,
                border_color="#555555"
            )
            content_frame.pack(fill="both", expand=True, padx=2, pady=2)
            
            # Add title label
            title_label = ctk.CTkLabel(
                content_frame, 
                text=f"Title: {title}", 
                font=("Segoe UI", 11, "bold"),
                anchor="w"
            )
            title_label.pack(fill="x", padx=5, pady=(5, 2))
            
            # Add type label
            type_label = ctk.CTkLabel(
                content_frame, 
                text=f"Type: {clip_type}", 
                font=("Segoe UI", 10),
                anchor="w"
            )
            type_label.pack(fill="x", padx=5, pady=(0, 5))
            
            # Show preview based on type
            if clip_type == "text" and isinstance(content, str):
                # For text content, show first few lines
                preview_text = content[:300] + ("..." if len(content) > 300 else "")
                text_area = Text(
                    content_frame,
                    wrap="word",
                    width=40,
                    height=6,
                    font=("Segoe UI", 10),
                    bg=MAIN_BG_COLOR,
                    fg=TEXT_FG_COLOR,
                    bd=0,
                    relief="flat",
                )
                text_area.insert("1.0", preview_text)
                text_area.configure(state="disabled")
                text_area.pack(fill="both", expand=True, padx=5, pady=5)
                
            elif clip_type == "image" and isinstance(content, bytes):
                # For image content, show a thumbnail
                try:
                    img = Image.open(BytesIO(content))
                    # Resize to fit preview
                    img.thumbnail((200, 200))
                    tk_img = ImageTk.PhotoImage(img)
                    # Need to keep a reference to prevent garbage collection
                    self.tk_image_references[id(self.preview_popup)] = tk_img
                    
                    img_label = Label(content_frame, image=tk_img, bg=MAIN_BG_COLOR)
                    img_label.pack(padx=5, pady=5)
                except Exception as e:
                    error_label = ctk.CTkLabel(
                        content_frame, 
                        text=f"[Image Preview Error: {str(e)}]",
                        text_color="red"
                    )
                    error_label.pack(padx=5, pady=5)
            
            # Size and position window
            self.preview_popup.update_idletasks()
            popup_width = self.preview_popup.winfo_reqwidth()
            popup_height = self.preview_popup.winfo_reqheight()
            
            # Adjust position to ensure popup is visible on screen
            screen_width = self.root.winfo_screenwidth()
            screen_height = self.root.winfo_screenheight()
            
            # Position 10px below the cursor
            x_pos = x
            y_pos = y + 10
            
            # Adjust if would go off screen
            if x_pos + popup_width > screen_width:
                x_pos = max(0, screen_width - popup_width - 10)
            if y_pos + popup_height > screen_height:
                y_pos = max(0, y - popup_height - 10)  # Show above instead of below
                
            self.preview_popup.geometry(f"{popup_width}x{popup_height}+{x_pos}+{y_pos}")
            
            # Add an escape key binding to close the popup
            self.preview_popup.bind("<Escape>", lambda e: self._hide_preview_popup())
            
        except Exception as e:
            print(f"Error showing preview popup: {e}")
            traceback.print_exc()
            self._hide_preview_popup()
            
    def _on_preview_destroy(self, event):
        """Handle preview popup destruction to clean up resources."""
        try:
            # Clean up the specific popup that was destroyed
            if hasattr(event, 'widget') and event.widget == self.preview_popup:
                popup_id = id(self.preview_popup)
                if popup_id in self.tk_image_references:
                    try:
                        del self.tk_image_references[popup_id]
                    except:
                        pass
                self.preview_popup = None
        except Exception as e:
            print(f"Warning: Error in preview destroy handler: {e}")

    def _hide_preview_popup(self):
        """Destroys the preview popup window and cleans up resources."""
        # Cancel any pending preview updates
        if hasattr(self, 'preview_update_id') and self.preview_update_id:
            try:
                self.root.after_cancel(self.preview_update_id)
                self.preview_update_id = None
            except:
                pass
            
        # Clean up the main preview popup
        if self.preview_popup:
            popup_id = id(self.preview_popup)
            try:
                self.preview_popup.destroy()
            except:
                pass
            if popup_id in self.tk_image_references:
                try:
                    del self.tk_image_references[popup_id]
                except:
                    pass
        self.preview_popup = None
        
        # Clean up widget tracking
        if hasattr(self, 'last_preview_widget'):
            self.last_preview_widget = None
        
        # Find and destroy all preview popups by title
        popups_to_destroy = []
        try:
            # First collect all popups to avoid modification during iteration
            for child in self.root.winfo_children():
                if (isinstance(child, ctk.CTkToplevel) and 
                    child.winfo_exists() and 
                    child != self.root and
                    hasattr(child, 'title') and
                    child.title() == "Clip Preview"):
                    popups_to_destroy.append(child)
                    
            # Now destroy them
            for popup in popups_to_destroy:
                try:
                    popup.destroy()
                except:
                    pass
        except Exception as e:
            print(f"Warning: Error cleaning up preview popups: {e}")

    def _select_clip_from_modal(self, history_index, modal):
        """Navigates main display to selected clip and closes modal."""
        self._hide_preview_popup()
        self._finalize_text_edit()
        try:
            if history_index in self.filtered_history_indices:
                self.current_filtered_index = self.filtered_history_indices.index(
                    history_index
                )
            else:
                self.search_var.set("")
                self._filter_history()
                self.current_filtered_index = (
                    self.filtered_history_indices.index(history_index)
                    if history_index in self.filtered_history_indices
                    else 0
                )
            self._show_clip()
            modal.destroy()
        except Exception as e:
            print(f"Error select modal: {e}")
            traceback.print_exc()
            try:
                modal.destroy()
            except:
                pass

    # WinCB-Elite_part5.py
    # Part 5 of 5 for WinCB-Elite.pyw
    # Contains batch save, save/restore group, window management, tray setup, and main execution
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into WinCB-Elite.pyw

    # --- Save Batch ---

    def _sanitize_filename(self, name):
        """Sanitizes a string for use as a filename."""
        name = re.sub(r'[<>:"/\\|?*]', "", name).strip(". ")
        if re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$", name, re.IGNORECASE):
            name = "_" + name
        return name if name else "Untitled_WinCB-Elite_Batch"

    def _prompt_and_save_batch(self):
        """Prompts for batch name and initiates save."""
        self._finalize_text_edit()
        if not self.running or not self.filtered_history_indices:
            self._show_popup("No clips visible.")
            return
        was_hidden = not self.root.winfo_viewable()
        if was_hidden:
            self._do_show_window()
        self.root.after(200, lambda wh=was_hidden: self._execute_save_batch(wh))

    def _execute_save_batch(self, was_hidden):
        """Gets input filename and saves filtered clips to file."""
        if not self.running:
            return
        dialog = ctk.CTkInputDialog(
            text="Enter name for batch export:", title="Export Data"
        )
        dialog.attributes("-topmost", True)
        name_raw = dialog.get_input()
        if not self.running or name_raw is None:
            if was_hidden and self.running:
                self.root.after(100, self._hide_window)
            return
        if not name_raw or not name_raw.strip():
            self._show_popup("Name empty.")
            return

        name = self._sanitize_filename(name_raw)
        default_filename = f"{name}.txt"
        
        # Use file dialog to let user choose save location
        try:
            from tkinter import filedialog
            dialog.attributes("-topmost", False)
            filepath = filedialog.asksaveasfilename(
                initialdir=str(BATCH_SAVE_DIR),
                initialfile=default_filename,
                title="Save Batch Export",
                defaultextension=".txt",
                filetypes=[("Text Files", "*.txt"), ("All Files", "*.*")]
            )
            dialog.attributes("-topmost", True)
        except Exception as file_dialog_err:
            print(f"Error showing file dialog: {file_dialog_err}")
            filepath = os.path.join(str(BATCH_SAVE_DIR), default_filename)
            
        if not filepath:
            if was_hidden and self.running:
                self.root.after(100, self._hide_window)
            return

        num_save = len(self.filtered_history_indices)
        texts = 0
        images = []
        try:
            search_text = self.search_entry.get()  # No placeholder check needed
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(
                    f"WinCB-Elite Export: {name}\nFilter: '{search_text}'\nClips: {num_save}\n{'='*40}\n\n"
                )
                for filt_idx, orig_idx in enumerate(self.filtered_history_indices):
                    if 0 <= orig_idx < len(self.history):
                        item = self.history[orig_idx]
                        title = item.get("title", "?")
                        kind = item.get("type")
                        content = item.get("content")
                        tags = item.get("tags", [])
                        
                        f.write(
                            f"--- Clip {filt_idx+1}/{num_save} ---\nTitle: {title}\nType: {kind}\n"
                        )
                        
                        # Add tags to the export if they exist
                        if tags:
                            tag_colors = [f"{tag} ({self._get_tag_color(tag)})" for tag in tags]
                            f.write(f"Tags: {', '.join(tag_colors)}\n")
                        # Handle text content
                        if kind == "text" and isinstance(content, str):
                            f.write(f"Content:\n{content}\n{'-'*20}\n\n")
                            texts += 1
                        # Handle image content
                        elif kind == "image":
                            # For images, include both [Image] indicator and the title as the image reference
                            image_info = f"[Image: {title}]"
                            if isinstance(content, bytes):
                                image_size = len(content)
                                image_info += f" (Size: {image_size/1024:.1f} KB)"
                            f.write(f"Content: {image_info}\n{'-'*20}\n\n")
                            images.append(title)
                        # Handle mixed content (containing both text and image)
                        elif kind == "mixed" and isinstance(content, dict):
                            text_part = content.get("text", "")
                            image_data = content.get("image")
                            
                            # Write text part if available
                            if text_part:
                                f.write(f"Content (Text):\n{text_part}\n")
                                texts += 1
                            
                            # Write image part if available
                            if isinstance(image_data, bytes):
                                image_size = len(image_data)
                                f.write(f"Content (Image): [Image: {title}] (Size: {image_size/1024:.1f} KB)\n")
                                images.append(title)
                                
                            f.write(f"{'-'*20}\n\n")
                        else:
                            f.write(f"Content: [Unknown format]\n{'-'*20}\n\n")
                    else:
                        f.write(
                            f"--- Clip {filt_idx+1}/{num_save} ---\nError: Bad Index\n{'-'*20}\n\n"
                        )
            
            # Extract just the filename for the message
            filename = os.path.basename(filepath)
            export_dir = os.path.dirname(filepath)
            
            msg = (
                f"Saved {texts} text(s)"
                + (f" and {len(images)} image title(s)" if images else "")
                + f" to:\n{filename}"
            )
            self._show_popup(msg)
            try:
                os.startfile(export_dir)
            except Exception as open_e:
                print(f"Warning: Failed to open export folder: {open_e}")
        except Exception as e:
            print(f"Error saving batch: {e}")
            traceback.print_exc()
            self._show_popup(f"Error saving: {e}")
        finally:
            if was_hidden and self.running:
                self.root.after(100, self._hide_window)

    def _prompt_save_restore_group(self):
        """Prompts user to save the current clip group or restore a clip group from a JSON file."""
        self._finalize_text_edit()
        if not self.running:
            return
        was_hidden = not self.root.winfo_viewable()
        if was_hidden:
            self._do_show_window()
        self.root.after(200, lambda wh=was_hidden: self._execute_save_restore_group(wh))

    def _execute_save_restore_group(self, was_hidden):
        """Executes the save/restore clip group dialog."""
        if not self.running:
            return
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Save/Restore Clip Group")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)

        ctk.CTkLabel(
            dialog, text="Choose an action:", font=("Segoe UI", 12, "bold")
        ).pack(pady=10, padx=10)

        def save_group():
            try:
                # First pop up a clean dialog to get the clip group name
                name_dialog = ctk.CTkToplevel(dialog)
                name_dialog.title("Name Clip Group")
                name_dialog.geometry("400x200")
                name_dialog.transient(dialog)
                name_dialog.grab_set()
                name_dialog.attributes("-topmost", True)
                name_dialog.configure(fg_color=MAIN_BG_COLOR)
                
                # Center the name dialog
                self._center_toplevel(name_dialog)
                
                # Add name input field with label
                ctk.CTkLabel(
                    name_dialog, 
                    text="Enter a name for this clip group:", 
                    font=("Segoe UI", 12, "bold")
                ).pack(pady=(20, 10), padx=20)
                
                # Pre-populate with current group name if one exists
                initial_name = self.current_group_name if self.current_group_name else ""
                name_var = ctk.StringVar(value=initial_name)
                name_entry = ctk.CTkEntry(
                    name_dialog,
                    textvariable=name_var,
                    width=300,
                    height=35,
                    placeholder_text="My Clip Group"
                )
                name_entry.pack(pady=10, padx=20)
                name_entry.focus_set()  # Auto-focus the entry field
                
                # If we have a pre-populated name, select all text for easy replacement
                if initial_name:
                    name_entry.select_range(0, "end")
                    # Schedule another selection to ensure it works
                    self.root.after(50, lambda: name_entry.select_range(0, "end"))
                
                # Ensure focus stays on the entry field with a slight delay
                self.root.after(100, lambda: name_entry.focus_set())
                
                # Variable to store the result
                result = {"name": None, "proceed": False}
                
                def on_name_ok():
                    group_name = name_var.get().strip()
                    if not group_name:
                        self._show_popup("Please enter a name for the clip group")
                        return
                    
                    # Validate the characters in the name
                    import re
                    valid_pattern = re.compile(r'^[A-Za-z0-9_\-.()\[\]]+$')
                    
                    if not valid_pattern.match(group_name):
                        # Create custom validation error popup
                        error_dialog = ctk.CTkToplevel(name_dialog)
                        error_dialog.title("Invalid Naming Convention")
                        error_dialog.geometry("400x250")
                        error_dialog.transient(name_dialog)
                        error_dialog.grab_set()
                        error_dialog.attributes("-topmost", True)
                        error_dialog.configure(fg_color=MAIN_BG_COLOR)
                        
                        # Center error dialog
                        self._center_toplevel(error_dialog)
                        
                        # Error message with requested format
                        ctk.CTkLabel(
                            error_dialog, 
                            text="Please only include the following characters for a valid naming convention:",
                            font=("Segoe UI", 12, "bold"),
                            wraplength=350
                        ).pack(pady=(20, 10), padx=20)
                        
                        # List of valid characters
                        char_frame = ctk.CTkFrame(error_dialog, fg_color="transparent")
                        char_frame.pack(pady=5, padx=20, fill="both", expand=True)
                        
                        ctk.CTkLabel(
                            char_frame,
                            text="• A-Z\n• a-z\n• 0-9\n• _ (underscore)\n• - (hyphen)\n• . (dot)\n• () (parentheses)\n• [] (square brackets)",
                            font=("Segoe UI", 11),
                            justify="left"
                        ).pack(pady=0, padx=20, anchor="w")
                        
                        # OK button
                        def on_error_ok():
                            error_dialog.destroy()
                            # Return focus to the name entry field
                            name_entry.focus_set()
                            # Select all text so user can easily delete or replace it
                            name_entry.select_range(0, "end")
                            # Schedule another selection just in case the first one doesn't stick
                            self.root.after(50, lambda: name_entry.select_range(0, "end"))
                            
                        ctk.CTkButton(
                            error_dialog,
                            text="OK",
                            width=100,
                            command=on_error_ok
                        ).pack(pady=15)
                        
                        # Wait for error dialog to close
                        name_dialog.wait_window(error_dialog)
                        return
                        
                    # Check if name ends with a dot
                    if group_name.endswith('.'):
                        self._show_popup("Filename cannot end with a dot (.)")
                        return
                    
                    # Save the name and mark to proceed
                    result["name"] = group_name
                    result["proceed"] = True
                    name_dialog.destroy()
                
                def on_name_cancel():
                    name_dialog.destroy()
                
                # Buttons for OK/Cancel
                btn_frame = ctk.CTkFrame(name_dialog, fg_color="transparent")
                btn_frame.pack(pady=20)
                
                # Create OK button with reference for tab order
                ok_button = ctk.CTkButton(
                    btn_frame, 
                    text="OK", 
                    width=120, 
                    command=on_name_ok
                )
                ok_button.pack(side="left", padx=10)
                
                # Create Cancel button
                cancel_button = ctk.CTkButton(
                    btn_frame, 
                    text="Cancel", 
                    width=120, 
                    command=on_name_cancel
                )
                cancel_button.pack(side="left", padx=10)
                
                # Set up explicit tab order using Tkinter's built-in mechanism
                name_entry.lift()  # First in tab order
                ok_button.lift()   # Second in tab order
                cancel_button.lift()  # Last in tab order
                
                # Make Enter key directly trigger OK from the entry field
                def on_enter(event):
                    print("Enter key pressed in name field - triggering OK")
                    on_name_ok()
                    return "break"
                
                # Bind Enter key in the entry to directly trigger OK
                name_entry.bind("<Return>", on_enter)
                name_entry.bind("<KP_Enter>", on_enter)  # Numpad Enter key
                
                # Provide a visible note about Enter key
                note_label = ctk.CTkLabel(
                    name_dialog,
                    text="Press Enter to confirm",
                    font=("Segoe UI", 10, "italic"),
                    text_color="#888888"  # Gray
                )
                note_label.pack(pady=(0, 10))
                
                # Wait for the dialog to be closed
                dialog.wait_window(name_dialog)
                
                # Check if user proceeded
                if not result["proceed"] or not result["name"]:
                    return
                
                # Now that we have the name, proceed with file selection
                from tkinter import filedialog
                
                # Sanitize filename
                group_name = self._sanitize_filename(result["name"])
                default_filename = f"{group_name}.json"

                dialog.attributes("-topmost", False)
                filename = filedialog.asksaveasfilename(
                    initialdir=str(BATCH_SAVE_DIR),
                    initialfile=default_filename,
                    title="Save Clip Group",
                    defaultextension=".json",
                    filetypes=[("JSON Files", "*.json")],
                )
                dialog.attributes("-topmost", True)
                
                if filename:
                    saveable_history = []
                    for item in self.history[:HISTORY_LIMIT]:
                        save_item = {
                            "type": item.get("type"),
                            "title": item.get("title", ""),
                            "timestamp": item.get("timestamp", 0),
                            "tags": item.get("tags", []),  # Add tags to saved data
                        }
                        content = item.get("content")
                        if save_item["type"] == "image" and isinstance(content, bytes):
                            save_item["content"] = base64.b64encode(content).decode(
                                "utf-8"
                            )
                        elif save_item["type"] == "text" and isinstance(content, str):
                            save_item["content"] = content
                        else:
                            print(
                                f"Warning: Skipping invalid item during group save: Type={save_item['type']}, Title='{save_item['title']}'"
                            )
                            continue
                        if save_item["type"] and save_item["content"] is not None:
                            saveable_history.append(save_item)
                    with open(filename, "w", encoding="utf-8") as f:
                        json.dump(saveable_history, f, ensure_ascii=False, indent=2)
                    self._show_popup(f"Saved clip group to:\n{filename}")
                    # Open the batchoutputs directory instead of the main app directory
                    os.startfile(str(BATCH_SAVE_DIR))
            except Exception as e:
                self._show_popup(f"Error saving clip group: {e}")
                print(f"Error saving clip group: {e}")
                traceback.print_exc()
            finally:
                try:
                    dialog.destroy()
                except:
                    pass
                if was_hidden and self.running:
                    self.root.after(100, self._hide_window)

        def restore_group():
            try:
                from tkinter import filedialog

                dialog.attributes("-topmost", False)
                filename = filedialog.askopenfilename(
                    initialdir=str(BATCH_SAVE_DIR),
                    title="Restore Clip Group",
                    filetypes=[("JSON Files", "*.json")],
                )
                dialog.attributes("-topmost", True)
                if filename:
                    with open(filename, "r", encoding="utf-8") as f:
                        saved_data = json.load(f)
                    temp_history = []
                    for item_data in saved_data:
                        if not isinstance(item_data, dict) or not all(
                            k in item_data for k in ("type", "content", "title")
                        ):
                            print(f"Skipping invalid history item: {item_data}")
                            continue
                        entry = {
                            "type": item_data["type"],
                            "title": item_data.get("title", ""),
                            "timestamp": item_data.get("timestamp", 0),
                            "tags": item_data.get("tags", []),  # Add tags to saved data
                        }
                        content = item_data["content"]
                        if entry["type"] == "image":
                            try:
                                entry["content"] = base64.b64decode(content)
                            except Exception as e:
                                print(
                                    f"Skipping invalid image data '{entry['title']}': {e}"
                                )
                                continue
                        elif entry["type"] == "text":
                            if isinstance(content, str):
                                entry["content"] = content
                            else:
                                print(f"Skipping invalid text entry '{entry['title']}'")
                                continue
                        else:
                            print(f"Skipping unknown type: {entry['type']}")
                            continue
                        temp_history.append(entry)
                    self.history = temp_history[:HISTORY_LIMIT]
                    
                    # Extract group name from filename
                    try:
                        base_name = os.path.basename(filename)
                        group_name = os.path.splitext(base_name)[0]  # Remove extension
                        self.current_group_name = group_name
                        self._update_group_display()
                        print(f"Loaded clip group: {group_name}")
                    except Exception as name_err:
                        print(f"Error extracting group name: {name_err}")
                        self.current_group_name = "Unknown Group"
                        self._update_group_display()
                    
                    self._filter_and_show()
                    self._show_popup(f"Restored clip group: {self.current_group_name}\nFrom: {filename}")
            except Exception as e:
                self._show_popup(f"Error restoring clip group: {e}")
                print(f"Error restoring clip group: {e}")
                traceback.print_exc()
            finally:
                try:
                    dialog.destroy()
                except:
                    pass
                if was_hidden and self.running:
                    self.root.after(100, self._hide_window)

        def cancel():
            try:
                dialog.destroy()
            except:
                pass
            if was_hidden and self.running:
                self.root.after(100, self._hide_window)

        btn_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        btn_frame.pack(pady=20)
        ctk.CTkButton(btn_frame, text="Save Group", width=120, command=save_group).pack(
            side="left", padx=5
        )
        ctk.CTkButton(
            btn_frame, text="Restore Group", width=120, command=restore_group
        ).pack(side="left", padx=5)
        ctk.CTkButton(btn_frame, text="Cancel", width=120, command=cancel).pack(
            side="left", padx=5
        )

        self._center_toplevel(dialog)

    # --- Scrolling and Window Management ---

    def _on_scrollable_configure(self, e=None):
        """Updates scrollregion with a shorter delay."""
        self.root.after(10, self._update_scrollregion)  # Reduced delay to 10ms

    def _update_scrollregion(self):
        """Sets the scrollable region based on content."""
        if (
            not self.running
            or not hasattr(self, "scrollable")
            or not self.scrollable.winfo_exists()
        ):
            return
        if (
            not hasattr(self.scrollable, "_parent_canvas")
            or not self.scrollable._parent_canvas.winfo_exists()
        ):
            return
        try:
            self.scrollable._parent_canvas.update_idletasks()
            bbox = self.scrollable._parent_canvas.bbox("all")
            if bbox:
                w = max(self.scrollable._parent_canvas.winfo_width(), bbox[2])
                h = bbox[3] + 10
                self.scrollable._parent_canvas.configure(scrollregion=(0, 0, w, h))
            else:
                self.scrollable._parent_canvas.configure(scrollregion=(0, 0, 1, 1))
        except Exception as e:
            print(f"Warn: Scroll update: {e}")

    def _hide_window(self):
        """Hides main window, shows tray icon."""
        if not self.running:
            return
        self._finalize_text_edit()
        self.root.withdraw()
        if hasattr(self, "icon") and self.icon and self.icon.HAS_NOTIFICATION:
            try:
                if not self.icon.visible:
                    self.icon.visible = True
                    print("Window hidden.")
            except Exception as e:
                print(f"Warn: Tray show err: {e}")

    def _show_window(self, *args):
        """Schedules showing the window."""
        if self.running:
            self.root.after(0, self._do_show_window)

    def _do_show_window(self):
        """Shows, lifts, focuses window, hides tray icon."""
        if not self.running:
            return
        try:
            self.root.deiconify()
            self.root.lift()
            self.root.attributes("-topmost", True)
            self.root.focus_force()
            self.root.after(
                500,
                lambda: (
                    self.root.attributes("-topmost", False) if self.running else None
                ),
            )
            if hasattr(self, "icon") and self.icon and self.icon.HAS_NOTIFICATION:
                try:
                    if self.icon.visible:
                        self.icon.visible = False
                        print("Window shown.")
                except Exception as e:
                    print(f"Warn: Tray hide err: {e}")
        except Exception as e:
            print(f"Error show window: {e}")

    def _center_toplevel(self, toplevel):
        """Centers a Toplevel window relative to the main root."""
        try:
            toplevel.update_idletasks()
            w, h = toplevel.winfo_width(), toplevel.winfo_height()
            if self.root.winfo_viewable():
                px, py, pw, ph = (
                    self.root.winfo_x(),
                    self.root.winfo_y(),
                    self.root.winfo_width(),
                    self.root.winfo_height(),
                )
                x = px + (pw // 2) - (w // 2)
                y = py + (ph // 2) - (h // 2)
            else:
                sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
                x = (sw // 2) - (w // 2)
                y = (sh // 2) - (h // 2)
            toplevel.geometry(f"+{max(0, x)}+{max(0, y)}")
            toplevel.lift()
        except Exception as e:
            print(f"Warn: Center Toplevel fail: {e}")

    def _confirm_action(self, title, message, action_callback):
        """Shows a confirmation dialog and calls action_callback if confirmed.
        
        A centralized way to confirm destructive actions.
        """
        # Create confirmation dialog
        dialog = ctk.CTkToplevel(self.root)
        dialog.title(title)
        dialog.geometry("400x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)
        
        # Message
        ctk.CTkLabel(
            dialog, 
            text=message,
            wraplength=350,
            justify="center"
        ).pack(pady=(20, 15), padx=20)
        
        # Buttons
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(pady=(0, 15))
        
        def on_cancel():
            dialog.destroy()
            
        def on_confirm():
            dialog.destroy()
            # Call the callback if provided
            if action_callback and callable(action_callback):
                action_callback()
        
        ctk.CTkButton(
            buttons_frame,
            text="Cancel",
            width=100,
            command=on_cancel
        ).pack(side="left", padx=10)
        
        ctk.CTkButton(
            buttons_frame,
            text="Yes, Proceed",
            width=100,
            fg_color="#8B0000",  # Dark red
            hover_color="#A00000",  # Slightly lighter red
            command=on_confirm
        ).pack(side="left", padx=10)
        
        # Center the dialog
        self._center_toplevel(dialog)

    # --- Close/Quit ---

    def _prompt_close(self):
        """Shows close confirmation dialog."""
        self._finalize_text_edit()
        if not self.running:
            return
        was_hidden = not self.root.winfo_viewable()
        if was_hidden:
            self._do_show_window()
        self.root.after(150, lambda wh=was_hidden: self._execute_close_prompt(wh))

    def _execute_close_prompt(self, was_hidden):
        """Displays the actual close confirmation dialog."""
        if not self.running:
            return
        try:
            log_dir = str(APP_DATA_DIR)
        except:
            log_dir = "current directory"
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Close WinCB-Elite")
        dialog.geometry("350x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)
        ctk.CTkLabel(
            dialog, text=f"History file location:\n{log_dir}\n\nClose WinCB-Elite now?"
        ).pack(pady=20, padx=10)
        btns = ctk.CTkFrame(dialog, fg_color="transparent")
        btns.pack(pady=10)

        def on_yes():
            print("Close confirmed.")
            try:
                dialog.destroy()
            except:
                pass
            self.root.after(50, self._quit)

        def on_no():
            try:
                dialog.destroy()
            except:
                pass
            if was_hidden and self.running:
                self.root.after(100, self._hide_window)

        ctk.CTkButton(
            btns,
            text="Yes, Close",
            width=110,
            command=on_yes,
            fg_color="#8B0000",
            hover_color="#A00000",
        ).pack(side="left", padx=10)
        ctk.CTkButton(btns, text="No", width=110, command=on_no).pack(
            side="left", padx=10
        )
        self._center_toplevel(dialog)

    def _quit(self):
        """Shuts down the application gracefully."""
        if not self.running:
            return
        print("Quit requested...")
        self.running = False
        print("- Saving...")
        self._finalize_text_edit()
        
        # Ensure auto-pause settings are saved
        print(f"- Saving auto-pause settings: {self.auto_pause_seconds} seconds")
        self.config["auto_pause_seconds"] = self.auto_pause_seconds
        self._save_config()
        
        # Check if we have an active group loaded
        active_group_saved = False
        if hasattr(self, "current_group_name") and self.current_group_name:
            try:
                print(f"- Auto-saving current group: {self.current_group_name}")
                # Prepare file path using the current group name
                group_filename = f"{self.current_group_name}.json"
                target_path = os.path.join(str(BATCH_SAVE_DIR), group_filename)
                
                # Prepare history data for saving
                saveable_history = []
                for item in self.history[:HISTORY_LIMIT]:
                    save_item = {
                        "type": item.get("type"),
                        "title": item.get("title", ""),
                        "timestamp": item.get("timestamp", 0),
                        "tags": item.get("tags", []),  # Add tags to saved data
                    }
                    content = item.get("content")
                    if save_item["type"] == "image" and isinstance(content, bytes):
                        save_item["content"] = base64.b64encode(content).decode("utf-8")
                    elif save_item["type"] == "text" and isinstance(content, str):
                        save_item["content"] = content
                    else:
                        print(f"Warning: Skipping invalid item during group auto-save: Type={save_item['type']}, Title='{save_item['title']}'")
                        continue
                    
                    if save_item["type"] and save_item["content"] is not None:
                        saveable_history.append(save_item)
                
                # Create the directory if it doesn't exist
                os.makedirs(os.path.dirname(target_path), exist_ok=True)
                
                # Save the data
                with open(target_path, "w", encoding="utf-8") as f:
                    json.dump(saveable_history, f, ensure_ascii=False, indent=2)
                
                print(f"- Group '{self.current_group_name}' auto-saved successfully")
                active_group_saved = True
            except Exception as e:
                print(f"Warning: Failed to auto-save group {self.current_group_name}: {e}")
                traceback.print_exc()
        
        # Always save main history file, too
        self._save_history()
        
        if hasattr(self, "icon") and self.icon:
            print("- Stopping tray...")
            try:
                self.icon.stop()
            except Exception as e:
                print(f"  Warn: Tray stop: {e}")
        if hasattr(self, "icon_thread") and self.icon_thread.is_alive():
            self.icon_thread.join(timeout=0.5)
        print("- Cleaning up...")
        
        # Cancel auto-pause timer if active
        if hasattr(self, "auto_pause_timer") and self.auto_pause_timer:
            try:
                self.root.after_cancel(self.auto_pause_timer)
                self.auto_pause_timer = None
                print("- Auto-pause timer canceled")
            except Exception as timer_err:
                print(f"  Warn: Auto-pause timer cleanup error: {timer_err}")
                
        try:
            self._hide_preview_popup()
        except:
            pass
        self.tk_image_references.clear()
        self.current_display_image = None
        print("- Destroying window...")
        try:
            if hasattr(self, "root") and self.root.winfo_exists():
                self.root.destroy()
        except Exception as e:
            print(f"  Warn: Destroy err: {e}")
        
        if active_group_saved:
            print(f"WinCB-Elite shutdown complete (Group '{self.current_group_name}' auto-saved)")
        else:
            print("WinCB-Elite shutdown complete.")

    # --- Window Snapping ---

    def _snap_window(self, corner):
        """Ensures window is visible and schedules coordinate setting."""
        if not self.running:
            return
        self._finalize_text_edit()
        if not self.root.winfo_viewable():
            self._do_show_window()
        self.root.after(100, lambda c=corner: self._snap_window_coords(c))

    def _snap_window_coords(self, corner):
        """Applies geometry for snapping."""
        if not self.running or not self.root.winfo_exists():
            return
        try:
            self.root.update_idletasks()
            sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
            w, h = self.root.winfo_width(), self.root.winfo_height()
            if w < 100 or h < 100:
                w, h = 750, 550
            x, y, off = 0, 0, 40
            if corner == "top_right":
                x = sw - w
            elif corner == "bottom_left":
                y = sh - h - off
            elif corner == "bottom_right":
                x = sw - w
                y = sh - h - off
            elif corner == "center":
                x = (sw // 2) - (w // 2)
                y = (sh // 2) - (h // 2)
            self.root.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")
        except Exception as e:
            print(f"Warn: Snap error: {e}")

    def snap_top_left(self):
        """Snap window to top-left corner."""
        self._snap_window("top_left")

    def snap_top_right(self):
        """Snap window to top-right corner."""
        self._snap_window("top_right")

    def snap_bottom_left(self):
        """Snap window to bottom-left corner."""
        self._snap_window("bottom_left")

    def snap_bottom_right(self):
        """Snap window to bottom-right corner."""
        self._snap_window("bottom_right")

    def snap_center(self):
        """Snap window to center of screen."""
        self._snap_window("center")

    # --- System Tray Setup ---

    def _setup_tray(self):
        """Sets up the pystray Icon in a separate thread."""
        try:
            icon_image = self._icon_image()
            if not icon_image:
                print("Error: Tray icon failed.")
                self.icon = None
                return

            def schedule(f):
                return lambda: self.root.after(0, f) if self.running else None

            is_vis = lambda i: self.running and self.root.winfo_viewable()
            is_hid = lambda i: self.running and not self.root.winfo_viewable()
            can_clr = lambda i: self.running and bool(self.history)
            can_sv = lambda i: self.running and bool(self.filtered_history_indices)
            is_p = lambda i: self.running and self.capture_paused
            is_r = lambda i: self.running and not self.capture_paused

            menu = (
                item(
                    "Show WinCB-Elite",
                    schedule(self._do_show_window),
                    default=True,
                    enabled=is_hid,
                ),
                item("Hide WinCB-Elite", schedule(self._hide_window), enabled=is_vis),
                item("Pause Capture", schedule(self._toggle_capture), enabled=is_r),
                item("Resume Capture", schedule(self._toggle_capture), enabled=is_p),
                item("Auto-Pause Settings", schedule(self._configure_auto_pause)),
                pystray.Menu.SEPARATOR,
                item(
                    "Snap Window",
                    pystray.Menu(
                        item("Top-Left", schedule(self.snap_top_left)),
                        item("Top-Right", schedule(self.snap_top_right)),
                        item("Bottom-Left", schedule(self.snap_bottom_left)),
                        item("Bottom-Right", schedule(self.snap_bottom_right)),
                        item("Center", schedule(self.snap_center)),
                    ),
                    enabled=is_vis,
                ),
                pystray.Menu.SEPARATOR,
                item(
                    "Export Data", schedule(self._prompt_and_save_batch), enabled=can_sv
                ),
                item("Clear History", schedule(self.clear_history), enabled=can_clr),
                pystray.Menu.SEPARATOR,
                item("Quit WinCB-Elite", schedule(self._prompt_close), enabled=self.running),
            )

            self.icon = pystray.Icon(
                APP_NAME, icon_image, f"{APP_NAME} - History", pystray.Menu(*menu)
            )
            print("Tray thread starting...")
            self.icon.run()
        except ImportError:
            print("\nWarn: Tray requires 'pystray' and 'Pillow'. Install with pip.")
            self.icon = None
        except Exception as e:
            print(f"Tray setup error: {e}")
            traceback.print_exc()
            self.icon = None
        finally:
            print("Tray thread finished.")
            self.icon = None

    def _icon_image(self):
        """Generates the tray icon image (requires Pillow)."""
        try:
            # Try to load custom icon first
            custom_icon_path = self.config.get("custom_icon_path")
            if custom_icon_path and os.path.exists(custom_icon_path):
                try:
                    # Load and return the custom icon
                    custom_img = Image.open(custom_icon_path)
                    # Resize to standard size if needed
                    custom_img = custom_img.resize((64, 64))
                    # If we have a transparent background, use it
                    if custom_img.mode != 'RGBA':
                        custom_img = custom_img.convert('RGBA')
                    print(f"Using custom icon: {custom_icon_path}")
                    return custom_img
                except Exception as icon_err:
                    print(f"Error loading custom icon, falling back to default: {icon_err}")
                    # Continue to default icon creation
            
            # Default icon creation
            W, H, P = 64, 64, 8
            CW, CH = 20, 12
            BGC = (0, 0, 0, 0)
            BOARD = MAIN_BG_COLOR
            OUT = "cyan"
            CLIP = "cyan"
            LINES = "white"
            img = Image.new("RGBA", (W, H), BGC)
            draw = ImageDraw.Draw(img)
            draw.rectangle((P, P + 4, W - P, H - P), fill=BOARD, outline=OUT, width=3)
            cx = W // 2 - CW // 2
            draw.rectangle((cx, P - CH + 4, cx + CW, P + 4), fill=CLIP)
            ays, aye, axo = P + 4, P + 10, 8
            draw.line((cx, ays, cx - axo, aye), fill=CLIP, width=4)
            draw.line((cx + CW, ays, cx + CW + axo, aye), fill=CLIP, width=4)
            ly, lxs, lxe, lg = P + 18, P + 8, W - P - 8, 8
            for _ in range(3):
                if ly < H - P - 4:
                    draw.line((lxs, ly, lxe, ly), fill=LINES, width=3)
                ly += lg
            return img
        except Exception as e:
            print(f"Icon creation err: {e}")
            return None

    # --- Utility: Show Message Box ---

    def _show_popup(self, msg, priority=1):
        """Schedules showing a message box in the main thread with priority level.
        Priority: 1 = normal, 2 = important (forced)
        """
        print(f"DEBUG: Show popup: '{msg}' (priority {priority})")
        if self.running:
            # For priority=2 messages, always show them
            # For priority=1 messages, don't show them if they're about in-progress clips
            if priority > 1 or "in-progress" not in msg.lower():
                self.root.after(0, lambda m=msg: self._display_messagebox(m))
            
    def _display_messagebox(self, msg):
        """Displays a custom message box using CTkToplevel."""
        if not self.running or not self.root.winfo_exists():
            return
        try:
            dialog = ctk.CTkToplevel(self.root)
            dialog.title(APP_NAME)
            dialog.geometry("350x150")
            dialog.transient(self.root)
            dialog.grab_set()
            dialog.attributes("-topmost", True)
            dialog.configure(fg_color=MAIN_BG_COLOR)
            ctk.CTkLabel(dialog, text=msg, wraplength=300).pack(
                pady=20, padx=10, expand=True, fill="both"
            )
            ctk.CTkButton(dialog, text="OK", width=100, command=dialog.destroy).pack(
                pady=10
            )
            self._center_toplevel(dialog)
        except Exception as e:
            print(f"DEBUG: Popup error: {e}")
            print(f"{APP_NAME} Msg: {msg}")

    def show_context_menu(self, event):
        """Displays the right-click context menu with context-aware options."""
        if hasattr(self, 'context_menu') and self.context_menu.winfo_exists():
            self.context_menu.destroy()
        self.context_menu = ctk.CTkFrame(self.root, fg_color="gray25")
        
        # Determine context (text or image)
        is_image_context = event.widget == self.img_label
        is_text_context = event.widget == self.textbox
        
        # Check if there's a text selection
        has_text_selection = False
        if is_text_context:
            try:
                sel_range = self.textbox.tag_ranges("sel")
                has_text_selection = bool(sel_range)
            except TclError:
                pass
        
        # Create common buttons
        self.context_menu.copy_btn = ctk.CTkButton(
            self.context_menu, text="Copy Active Clip", 
            command=self.copy_active_clip_to_buffer, width=220
        )
        
        # Create context-specific buttons
        if is_image_context:
            self.context_menu.copy_image_btn = ctk.CTkButton(
                self.context_menu, text="Copy Image to New Clip", 
                command=self.copy_image_to_new_clip, width=220
            )
        
        self.context_menu.paste_to_current_btn = ctk.CTkButton(
            self.context_menu, text="Paste to Current Clip", 
            command=self.paste_from_buffer_to_current_clip, width=220
        )
        
        self.context_menu.paste_btn = ctk.CTkButton(
            self.context_menu, text="Paste from Clipboard", 
            command=self.paste_from_buffer_to_in_progress_clip, width=220
        )
        
        self.context_menu.clear_btn = ctk.CTkButton(
            self.context_menu, text="Clear Clipboard", 
            command=self.clear_additional_buffer, width=220
        )
        
        # Pack buttons in context-specific order
        self.context_menu.copy_btn.pack(pady=2)
        if is_image_context:
            self.context_menu.copy_image_btn.pack(pady=2)
        self.context_menu.paste_to_current_btn.pack(pady=2)
        self.context_menu.paste_btn.pack(pady=2)
        self.context_menu.clear_btn.pack(pady=2)

        # Configure button states
        can_paste = self.additional_clipboard is not None
        has_current_clip = (0 <= self.current_filtered_index < len(self.filtered_history_indices))
        
        self.context_menu.paste_btn.configure(state="normal" if can_paste else "disabled")
        self.context_menu.paste_to_current_btn.configure(state="normal" if (can_paste and has_current_clip) else "disabled")

        # Handle click-outside binding
        if self.click_outside_handler_id:
            try:
                self.root.unbind("<Button-1>", self.click_outside_handler_id)
            except TclError:
                pass
            self.click_outside_handler_id = None

        # IMPROVED MENU POSITIONING
        # First update the menu to calculate its size
        self.context_menu.update_idletasks()

        # Get screen dimensions
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()

        # Get menu dimensions
        menu_width = 240  # Force a reasonable fixed width since reqwidth might not be accurate
        menu_height = 150  # Estimate of height for 4 buttons with padding

        # Get mouse position - use actual event coordinates
        x_pos = event.x_root
        y_pos = event.y_root

        # Center menu on cursor (offset to the left)
        x_pos = x_pos - (menu_width // 2)  # Center horizontally on cursor
        y_pos = y_pos - 10  # Place slightly above cursor

        # Adjust position if it would go off screen
        if x_pos + menu_width > screen_width:
            x_pos = screen_width - menu_width - 10
        if x_pos < 0:
            x_pos = 10
        
        if y_pos + menu_height > screen_height:
            y_pos = screen_height - menu_height - 10
        if y_pos < 0:
            y_pos = 10

        # Convert to window-relative coordinates if needed
        try:
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
            x_pos = max(0, x_pos - root_x)
            y_pos = max(0, y_pos - root_y)
        except Exception as e:
            print(f"Warning: Error calculating menu position: {e}")

        # Place menu at adjusted coordinates
        self.context_menu.place(x=x_pos, y=y_pos)
        self.context_menu.lift()

        def click_outside_handler(e):
            if not self.running or not hasattr(self, 'context_menu') or not self.context_menu.winfo_exists():
                if hasattr(self, 'click_outside_handler_id') and self.click_outside_handler_id:
                    try: self.root.unbind("<Button-1>", self.click_outside_handler_id)
                    except TclError: pass
                    self.click_outside_handler_id = None
                return

            widget_under_cursor = e.widget
            is_on_menu_or_child = False
            current_widget = widget_under_cursor
            while current_widget is not None:
                if current_widget == self.context_menu:
                    is_on_menu_or_child = True
                    break
                try:
                    current_widget = current_widget.master 
                except AttributeError:
                    break 
            
            if not is_on_menu_or_child:
                self.context_menu.place_forget()
                if hasattr(self, 'click_outside_handler_id') and self.click_outside_handler_id:
                    try: self.root.unbind("<Button-1>", self.click_outside_handler_id)
                    except TclError: pass
                    self.click_outside_handler_id = None
        
        # Delay the binding of the click_outside_handler
        # Store the binding ID on self so it can be unbound.
        self.root.after(50, lambda: setattr(self, 'click_outside_handler_id', self.root.bind("<Button-1>", click_outside_handler, add="+")))

    def copy_active_clip_to_buffer(self, event=None):
        """Copies the full content of the currently displayed clip to the additional buffer."""
        if not self.running or not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            print("No active clip to copy to buffer.")
            return

        try:
            original_index = self.filtered_history_indices[self.current_filtered_index]
            if not (0 <= original_index < len(self.history)):
                print("Error copying to buffer: Invalid clip index.")
                return

            item = self.history[original_index]
            clip_type = item.get("type")
            content = item.get("content")

            if content is None:
                print("Cannot copy to buffer: No content available in the active clip.")
                return

            self.additional_clipboard = {"type": clip_type, "content": content}
            self._update_buffer_status()  # Update the buffer status indicator
            self._show_popup(f"Active '{clip_type}' clip copied to buffer")  # Show a popup message
            print(f"Active '{clip_type}' clip content copied to additional buffer.")
            if self.context_menu.winfo_ismapped():
                self.context_menu.place_forget()
        except Exception as e:
            print(f"Error copying active clip to buffer: {e}")
            traceback.print_exc()

    def copy_focused_content_to_buffer(self, event=None):
        """Copies selected text, all text from textbox, or current image to the additional buffer."""
        if not self.running:
            return

        copied_something = False
        # Priority 1: Selected text in the textbox
        if self.textbox.winfo_ismapped() and self.textbox.cget("state") == "normal":
            try:
                has_selection = self.textbox.tag_ranges("sel")
                if has_selection:
                    selection = self.textbox.get("sel.first", "sel.last")
                    if selection and selection.strip():
                        # Reset clipboard internal state to ensure fresh selection copy
                        self.last_clip_text = None
                        self.last_clip_img_data = None
                        
                        self.additional_clipboard = {"type": "text", "content": selection}
                        self._update_buffer_status()  # Update buffer status indicator
                        self._show_popup(f"Selected text copied")  # Show popup
                        print(f"DEBUG: Selected text copied: '{selection[:30]}...'")
                        self._force_copy_to_clipboard("text", selection)  # Also copy to system clipboard
                        copied_something = True
                        if self.context_menu.winfo_ismapped():
                            self.context_menu.place_forget()
                        return  # Exit early after copying selection
            except TclError: # No selection
                pass
        
        # Priority 2: If nothing selected, but textbox is active and showing a text clip
        if not copied_something and self.textbox.winfo_ismapped() and self.textbox.cget("state") == "normal":
            full_text = self.textbox.get("1.0", "end-1c")
            if full_text and full_text.strip():
                self.additional_clipboard = {"type": "text", "content": full_text}
                self._update_buffer_status()  # Update buffer status indicator
                self._show_popup(f"Full text copied to buffer")  # Show popup
                print("All text from textbox copied to additional buffer.")
                copied_something = True

        # Priority 3: Current image if displayed
        elif not copied_something and self.img_label.winfo_ismapped() and self.current_display_image:
            if (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
                try:
                    original_index = self.filtered_history_indices[self.current_filtered_index]
                    if (0 <= original_index < len(self.history) and 
                        self.history[original_index]["type"] == "image"):
                        image_content = self.history[original_index]["content"]
                        self.additional_clipboard = {"type": "image", "content": image_content}
                        self._update_buffer_status()  # Update buffer status indicator
                        self._show_popup(f"Image copied to buffer")  # Show popup
                        print("Current image copied to additional buffer.")
                        copied_something = True
                except Exception as e:
                    print(f"Error accessing image content for additional buffer: {e}")
            
        if not copied_something:
            print("No suitable content found to copy to additional buffer.")
        
        if self.context_menu.winfo_ismapped():
            self.context_menu.place_forget()

    def clear_additional_buffer(self, event=None):
        """Clears the additional clipboard buffer."""
        self.additional_clipboard = None
        self._update_buffer_status()  # Update the buffer status indicator
        self._show_popup("Buffer cleared")  # Show popup
        print("Additional buffer cleared.")
        if self.context_menu.winfo_ismapped():
            self.context_menu.place_forget()
            
    def paste_from_buffer_to_in_progress_clip(self, event=None):
        """Pastes content from the additional buffer directly to system clipboard."""
        print("DEBUG: paste_from_buffer_to_in_progress_clip called")
        
        if not self.additional_clipboard:
            print("DEBUG: Buffer is empty")
            self._show_popup("Buffer is empty", priority=2)  # Higher priority
            print("Additional buffer is empty. Nothing to paste.")
            return

        item_type = self.additional_clipboard["type"]
        item_content = self.additional_clipboard["content"]

        # First update the in-progress clip in the background
        print("DEBUG: Updating in-progress clip")
        in_progress_updated = False
        
        if item_type == "text":
            if self.in_progress_clip["text"]: # Add a newline if there's existing text
                self.in_progress_clip["text"] += "\n" + item_content
            else:
                self.in_progress_clip["text"] = item_content
            print(f"DEBUG: Text from buffer added to in-progress clip ({len(item_content)} chars)")
            in_progress_updated = True
            
        elif item_type == "image":
            self.in_progress_clip["images"].append(item_content)
            print(f"DEBUG: Image from buffer added to in-progress clip ({len(item_content)} bytes)")
            in_progress_updated = True
            
        elif item_type == "mixed" and isinstance(item_content, dict):
            # Handle mixed content by adding each component
            if "text" in item_content and item_content["text"]:
                if self.in_progress_clip["text"]:
                    self.in_progress_clip["text"] += "\n" + item_content["text"]
                else:
                    self.in_progress_clip["text"] = item_content["text"]
                print(f"DEBUG: Text from mixed buffer added to in-progress clip")
                in_progress_updated = True
            
            if "image" in item_content and item_content["image"]:
                self.in_progress_clip["images"].append(item_content["image"])
                print(f"DEBUG: Image from mixed buffer added to in-progress clip")
                in_progress_updated = True
        else:
            print(f"DEBUG: Unknown type in additional buffer: {item_type}")

        # Then try to copy to system clipboard (even if in-progress update failed)
        print("DEBUG: Copying buffer content to system clipboard")
        success = self._force_copy_to_clipboard(item_type, item_content)
        
        # Update feedback message based on both operations
        if success:
            msg = f"Content pasted to system clipboard"
            if in_progress_updated:
                if item_type == "text":
                    msg += " and added to in-progress clip"
                elif item_type == "image":
                    msg += " and image added to in-progress clip"
                elif item_type == "mixed":
                    msg += " and content added to in-progress clip"
                    
            self._show_popup(msg, priority=2)  # Higher priority
            print(f"DEBUG: Content copied to system clipboard: type={item_type}")
            
            # Update buffer status indicator
            self._update_buffer_status()
        else:
            self._show_popup("Failed to copy to clipboard. See console for details.", priority=2)
            print("DEBUG-ERROR: Failed to paste content to clipboard after retries")
            
        if self.context_menu.winfo_ismapped():
            self.context_menu.place_forget()

    def _update_buffer_status(self):
        """Updates the buffer status variable (but no longer shows in UI)."""
        if not self.additional_clipboard:
            self.buffer_status_var.set("Buffer: Empty")
            return
            
        buffer_type = self.additional_clipboard.get("type", "unknown")
        if buffer_type == "text":
            content = self.additional_clipboard.get("content", "")
            if content:
                preview = content[:20].replace("\n", " ")
                if len(content) > 20:
                    preview += "..."
                self.buffer_status_var.set(f"Buffer: Text \"{preview}\"")
            else:
                self.buffer_status_var.set("Buffer: Empty text")
        elif buffer_type == "image":
            self.buffer_status_var.set("Buffer: Image data")
        else:
            self.buffer_status_var.set(f"Buffer: {buffer_type.capitalize()}")
            
        # We keep this method for internal state tracking, but it no longer updates UI

    def start_new_clip_from_selection(self):
        """Clears the in-progress clip and adds current selection to it."""
        self.in_progress_clip = {"text": "", "images": []}
        print("New in-progress clip started.")
        self.add_selection_to_in_progress_clip(is_new=True)

    def add_selection_to_in_progress_clip(self, is_new=False):
        """Adds selected text or current image to the in-progress clip."""
        added_something = False
        if self.textbox.winfo_ismapped() and self.textbox.cget("state") == "normal":
            try:
                selection = self.textbox.get("sel.first", "sel.last")
                if selection and selection.strip():
                    if self.in_progress_clip["text"] and not is_new:
                         self.in_progress_clip["text"] += "\n" + selection
                    else:
                         self.in_progress_clip["text"] = selection
                    print(f"Selected text added to in-progress clip.")
                    added_something = True
            except TclError: # No text selection
                # If it's a new clip and no text selection, try to add full text
                if is_new:
                    full_text = self.textbox.get("1.0", "end-1c")
                    if full_text and full_text.strip():
                        self.in_progress_clip["text"] = full_text
                        print(f"Full text from editor added to new in-progress clip.")
                        added_something = True

        if not added_something and self.img_label.winfo_ismapped() and self.current_display_image:
            if (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
                try:
                    original_index = self.filtered_history_indices[self.current_filtered_index]
                    if (0 <= original_index < len(self.history) and 
                        self.history[original_index]["type"] == "image"):
                        image_content = self.history[original_index]["content"]
                        self.in_progress_clip["images"].append(image_content)
                        print(f"Current image added to in-progress clip.")
                        added_something = True
                except Exception as e:
                    print(f"Error adding image to in-progress clip: {e}")
        
        if not added_something:
            print("No selection or image to add to in-progress clip.")
        else:
            # Provide feedback about the in-progress clip status
            status = f"In-progress clip: "
            if self.in_progress_clip["text"]:
                status += f"{len(self.in_progress_clip['text'].splitlines())} text line(s)"
            if self.in_progress_clip["images"]:
                if self.in_progress_clip["text"]: status += ", "
                status += f"{len(self.in_progress_clip['images'])} image(s)"
            print(status)

    def save_in_progress_clip(self):
        """Saves the in-progress clip (text and images) as a new history item."""
        text_content = self.in_progress_clip["text"].strip()
        images_content = self.in_progress_clip["images"]

        if not text_content and not images_content:
            self._show_popup("In-progress clip is empty. Nothing to save.")
            return

        current_time = time.time()
        
        # Handle different combinations of content
        if text_content and images_content:
            # We have both text and images - create a mixed type clip
            print(f"DEBUG: Creating mixed clip with text and {len(images_content)} image(s)")
            
            # Use the first image for the mixed clip
            mixed_content = {
                "text": text_content,
                "image": images_content[0]  # First image goes into mixed clip
            }
            
            # Add the mixed clip to history
            self._add_to_history("mixed", mixed_content, is_from_selection=True)
            
            # If there are additional images, add them as separate clips
            if len(images_content) > 1:
                for img_data in images_content[1:]:
                    self._add_to_history("image", img_data, is_from_selection=True)
                    print("Additional in-progress image saved as separate clip.")
                
                self._show_popup(f"In-progress content saved: 1 mixed clip and {len(images_content)-1} additional image(s).")
            else:
                self._show_popup("In-progress content saved as mixed clip.")
                
        elif text_content:
            # Text only
            self._add_to_history("text", text_content, is_from_selection=True)
            print("In-progress text saved to history.")
            self._show_popup("In-progress text saved to history.")
            
        elif images_content:
            # Images only
            for img_data in images_content:
                self._add_to_history("image", img_data, is_from_selection=True)
            print(f"{len(images_content)} in-progress image(s) saved to history.")
            self._show_popup(f"{len(images_content)} in-progress image(s) saved to history.")

        # Reset after saving
        self.in_progress_clip = {"text": "", "images": []}
        self._filter_and_show() # Refresh display

    def paste_from_buffer_to_current_clip(self, event=None):
        """Pastes content from the additional buffer into the current clip."""
        print("DEBUG: paste_from_buffer_to_current_clip called")
        
        if not self.additional_clipboard:
            print("DEBUG: Buffer is empty")
            self._show_popup("Buffer is empty")
            print("Buffer is empty. Cannot paste to current clip.")
            return
        
        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            print("DEBUG: No current clip selected")
            self._show_popup("No current clip selected")
            print("No current clip selected to paste into.")
            return
        
        print("DEBUG: Getting clip details")
        original_index = self.filtered_history_indices[self.current_filtered_index]
        current_clip = self.history[original_index]
        print(f"DEBUG: Current clip type: {current_clip['type']}")
        print(f"DEBUG: Buffer clip type: {self.additional_clipboard['type']}")
        
        # Handle text-to-text pasting - directly add to the text widget
        if self.additional_clipboard["type"] == "text" and current_clip["type"] == "text":
            print("DEBUG: Text-to-text pasting")
            content_to_paste = self.additional_clipboard["content"]
            self.textbox.configure(state="normal")
            self.textbox.insert("end", "\n" + content_to_paste)  # Append with newline
            self._on_text_edited()  # Trigger save
            self._show_popup("Text pasted to current clip")
            print("Pasted text content to current text clip.")
            return
            
        # Handle text-to-mixed pasting
        if self.additional_clipboard["type"] == "text" and current_clip["type"] == "mixed":
            print("DEBUG: Text-to-mixed pasting")
            content_to_paste = self.additional_clipboard["content"]
            
            # Get the current text content from the mixed clip
            current_text = current_clip["content"].get("text", "")
            
            # Append the new text with a newline
            new_text = current_text + "\n" + content_to_paste if current_text else content_to_paste
            
            # Update the mixed clip's text content
            current_clip["content"]["text"] = new_text
            
            # Update the display
            self.textbox.configure(state="normal")
            self.textbox.delete("1.0", "end")
            self.textbox.insert("1.0", new_text)
            
            # Save changes
            self.history[original_index]["timestamp"] = time.time()
            self._save_history()
            
            self._show_popup("Text pasted to current mixed clip")
            print("Pasted text content to current mixed clip.")
            return
        
        # Handle image-to-text pasting - convert text clip to mixed in place
        if self.additional_clipboard["type"] == "image" and current_clip["type"] == "text":
            print("DEBUG: Image-to-text pasting - converting to mixed in place")
            image_content = self.additional_clipboard["content"]
            text_content = current_clip["content"]
            
            # Convert the current text clip to a mixed clip in-place
            mixed_content = {
                "text": text_content,
                "image": image_content
            }
            
            # Update the current clip
            self.history[original_index]["type"] = "mixed"
            self.history[original_index]["content"] = mixed_content
            self.history[original_index]["timestamp"] = time.time()
            
            # Save changes
            self._save_history()
            
            # Refresh the display without creating a new clip
            self._show_clip()
            
            self._show_popup("Image added to current clip")
            print("Converted text clip to mixed clip with image and text")
            return
            
        # Handle image-to-mixed pasting - add image to the mixed content
        if self.additional_clipboard["type"] == "image" and current_clip["type"] == "mixed":
            print("DEBUG: Image-to-mixed pasting - updating mixed clip in place")
            image_content = self.additional_clipboard["content"]
            
            # Update the image in the mixed clip
            current_clip["content"]["image"] = image_content
            current_clip["timestamp"] = time.time()
            
            # Save changes
            self._save_history()
            
            # Refresh the display
            self._show_clip()
            
            self._show_popup("Image updated in current clip")
            print("Updated image in mixed clip")
            return
        
        # Handle image-to-image pasting - replace image with new one
        if self.additional_clipboard["type"] == "image" and current_clip["type"] == "image":
            print("DEBUG: Image-to-image pasting - replacing image")
            image_content = self.additional_clipboard["content"]
            
            # Replace the image content
            current_clip["content"] = image_content
            current_clip["timestamp"] = time.time()
            
            # Save changes
            self._save_history()
            
            # Refresh the display
            self._show_clip()
            
            self._show_popup("Image replaced in current clip")
            print("Replaced image in current clip")
            return
            
        # For other combinations, create a new clip
        buffer_type = self.additional_clipboard["type"]
        buffer_content = self.additional_clipboard["content"]
        
        # Create descriptive title for the new clip
        current_title = current_clip.get("title", "")
        
        print("DEBUG: Creating mixed/combined clip")
        if buffer_type == "text":
            new_title = f"Combined: {current_title}"
            # Create a mixed type content with both the current clip and buffer
            if current_clip["type"] == "image":
                # Combine image and text into mixed type
                print("DEBUG: Creating mixed type with image+text")
                # IMPORTANT: Structure mixed content as a proper dictionary with text and image keys
                mixed_content = {
                    "text": buffer_content,
                    "image": current_clip["content"]
                }
                
                # Update the current clip in-place instead of creating a new one
                self.history[original_index]["type"] = "mixed"
                self.history[original_index]["content"] = mixed_content
                self.history[original_index]["timestamp"] = time.time()
                
                # Save changes
                self._save_history()
                
                # Refresh the display
                self._show_clip()
                
                self._show_popup("Text added to current image clip")
                print("Converted image clip to mixed clip with text")
                return
                
            else:
                # Create combined text clip
                print("DEBUG: Creating combined text clip")
                # If current clip is already mixed, get its text content
                if current_clip["type"] == "mixed":
                    current_text = current_clip["content"].get("text", "")
                    combined_text = current_text + "\n" + buffer_content if current_text else buffer_content
                else:
                    combined_text = current_clip["content"] + "\n" + buffer_content
                
                # Update the current clip in-place
                if current_clip["type"] == "text":
                    current_clip["content"] = combined_text
                    current_clip["timestamp"] = time.time()
                    self._save_history()
                    
                    # Refresh text display
                    self.textbox.configure(state="normal")
                    self.textbox.delete("1.0", "end")
                    self.textbox.insert("1.0", combined_text)
                    
                    self._show_popup("Text added to current clip")
                    print("Added text to current text clip")
                    return
        
        # If we get here, we couldn't handle the combination in-place
        self._show_popup("Cannot combine these clip types directly")
        print("Cannot combine these clip types directly")

    def copy_image_to_new_clip(self, event=None):
        """Creates a new mixed-type clip from the current image and makes it editable."""
        print("DEBUG: copy_image_to_new_clip called")
        
        if not self.img_label.winfo_ismapped() or not self.current_display_image:
            print("DEBUG: No image to create clip from")
            print("No image to create clip from.")
            return
        
        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            print("DEBUG: No clip selected")
            print("No clip selected.")
            return
        
        try:
            print("DEBUG: Getting image clip details")
            original_index = self.filtered_history_indices[self.current_filtered_index]
            if not (0 <= original_index < len(self.history)):
                print("DEBUG: Invalid history index")
                print("Error: Invalid history index.")
                return
            
            current_clip = self.history[original_index]
            if current_clip["type"] != "image":
                print(f"DEBUG: Selected clip is not an image: {current_clip['type']}")
                print("Selected clip is not an image.")
                return
            
            # Create a mixed-type clip directly
            print("DEBUG: Creating mixed-type clip from image")
            image_content = current_clip["content"]
            print(f"DEBUG: Image content type: {type(image_content)}, size: {len(image_content)} bytes")
            
            # FIXED: Create proper dictionary structure for mixed content
            mixed_content = {
                "text": "",  # Start with empty text
                "image": image_content
            }
            
            print(f"DEBUG: Mixed content keys: {list(mixed_content.keys())}")
            print(f"DEBUG: Text type: {type(mixed_content['text'])}")
            print(f"DEBUG: Image type: {type(mixed_content['image'])}")
            
            # Add directly to history as a mixed type
            print("DEBUG: Adding mixed content to history")
            self._add_to_history("mixed", mixed_content, is_from_selection=True)
            
            # Navigate to the new clip
            print("DEBUG: Refreshing display to show new mixed clip")
            self._filter_history()
            self.current_filtered_index = 0
            self._show_clip()
            
            print("DEBUG: Image copied to new editable mixed-type clip")
            self._show_popup("Image copied to new editable clip")
        except Exception as e:
            print(f"DEBUG-ERROR: Error creating mixed clip: {e}")
            traceback.print_exc()
            print(f"Error creating mixed clip: {e}")
        
        if self.context_menu.winfo_ismapped():
            self.context_menu.place_forget()

    def _context_aware_copy(self, event):
        """Handles context-aware copy operations."""
        if event.widget == self.textbox:
            # Handle copy from text widget - only copy selected text if there is a selection
            try:
                has_selection = self.textbox.tag_ranges("sel")
                if has_selection:
                    # Copy only the selected text to the clipboard
                    selection = self.textbox.get("sel.first", "sel.last")
                    if selection and selection.strip():
                        # Copy selection to system clipboard
                        self._force_copy_to_clipboard("text", selection)
                        print(f"DEBUG: Selected text copied to clipboard: {len(selection)} chars")
                        return
            except TclError:
                pass  # No selection
            
            # If no selection, fall back to copying the focused content
            self.copy_focused_content_to_buffer()
        else:
            # Default to copying the entire active clip
            self.copy_clip_to_clipboard()

    def _context_aware_paste(self, event):
        """Handles context-aware paste operations."""
        if event.widget == self.textbox and self.textbox.cget("state") == "normal":
            # Handle paste into textbox - use the system clipboard directly
            try:
                win32clipboard.OpenClipboard()
                if win32clipboard.IsClipboardFormatAvailable(win32clipboard.CF_UNICODETEXT):
                    text = win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
                    win32clipboard.CloseClipboard()
                    
                    # Check if there's a selection to replace
                    try:
                        has_selection = self.textbox.tag_ranges("sel")
                        if has_selection:
                            # Replace selected text with clipboard content
                            self.textbox.delete("sel.first", "sel.last")
                    except TclError:
                        pass  # No selection to replace
                    
                    # Insert clipboard content at current position
                    self.textbox.insert("insert", text)
                    self._on_text_edited()  # Mark as modified to trigger save
                    print(f"DEBUG: Pasted {len(text)} chars at cursor position")
                    return
                win32clipboard.CloseClipboard()
            except Exception as e:
                print(f"DEBUG: Error during direct paste: {e}")
                try:
                    win32clipboard.CloseClipboard()
                except:
                    pass
        
        # Default to the standard buffer paste behavior
        self.paste_from_buffer_to_in_progress_clip()

    def _change_app_icon(self):
        """Opens a file dialog to select a new application icon and updates it immediately."""
        if not self.running:
            return
            
        try:
            from tkinter import filedialog
            
            # Get current icon path as initial directory
            current_path = self.config.get("custom_icon_path")
            initial_dir = os.path.dirname(current_path) if current_path else str(APP_DATA_DIR)
            
            # Show file dialog to select an image file
            icon_path = filedialog.askopenfilename(
                title="Select Application Icon",
                initialdir=initial_dir,
                filetypes=[
                    ("Image Files", "*.png *.jpg *.jpeg *.bmp *.ico *.gif"),
                    ("All Files", "*.*")
                ]
            )
            
            if not icon_path:  # User canceled
                return
                
            # Validate the selected file
            if not os.path.exists(icon_path):
                self._show_popup("Selected file does not exist.")
                return
                
            # Update config with new icon path
            self.config["custom_icon_path"] = icon_path
            self._save_config()
            
            # Force update window icon immediately
            try:
                # For Windows, we need to ensure the file is in .ico format
                # For non-ico images, convert them first
                if not icon_path.lower().endswith('.ico'):
                    # Create a temporary ico file
                    temp_ico = os.path.join(str(APP_DATA_DIR), "temp_icon.ico")
                    img = Image.open(icon_path)
                    if img.mode != 'RGBA':
                        img = img.convert('RGBA')
                    img.save(temp_ico, format="ICO", sizes=[(32, 32)])
                    icon_path_to_use = temp_ico
                else:
                    icon_path_to_use = icon_path
                
                # Force icon update
                self.root.iconbitmap(default=icon_path_to_use)
                
                # Also update using wm_iconbitmap
                self.root.wm_iconbitmap(icon_path_to_use)
                
                # Force window refresh to apply icon change
                self.root.update_idletasks()
                
                # For extra measure, re-set the title to trigger a window manager refresh
                current_title = self.root.title()
                self.root.title(current_title + " ")
                self.root.update_idletasks()
                self.root.title(current_title)
                
                # Windows-specific: Set taskbar icon using ctypes and win32 APIs
                try:
                    import ctypes
                    import win32gui
                    import win32con
                    import win32api
                    
                    # Get handle to the icon file
                    icon_handle = win32gui.LoadImage(
                        0, 
                        icon_path_to_use,
                        win32con.IMAGE_ICON,
                        0, 0,
                        win32con.LR_LOADFROMFILE
                    )
                    
                    # Get window handle
                    hwnd = ctypes.windll.user32.GetParent(self.root.winfo_id())
                    
                    # Set both small and large icons (0 = small, 1 = large/taskbar)
                    win32gui.SendMessage(hwnd, win32con.WM_SETICON, 0, icon_handle)
                    win32gui.SendMessage(hwnd, win32con.WM_SETICON, 1, icon_handle)
                    
                    # Force taskbar refresh
                    ctypes.windll.user32.FlashWindow(hwnd, 0)
                    
                    # Tell Windows to update the non-client area (window border, titlebar, etc.)
                    win32gui.SetWindowPos(
                        hwnd, 0, 0, 0, 0, 0, 
                        win32con.SWP_NOMOVE | win32con.SWP_NOSIZE | 
                        win32con.SWP_NOACTIVATE | win32con.SWP_NOZORDER | 
                        win32con.SWP_FRAMECHANGED
                    )
                    
                except Exception as taskbar_err:
                    print(f"Windows taskbar icon update failed: {taskbar_err}")
                
                print(f"Window icon updated to: {icon_path_to_use}")
            except Exception as window_icon_err:
                print(f"Could not update window icon: {window_icon_err}")
            
            # Apply the new icon to system tray
            self._reload_tray_icon()
            
            # Show confirmation
            self._show_popup(f"Icon changed to:\n{os.path.basename(icon_path)}", priority=2)
            
            # Log the change
            print(f"Custom icon set to: {icon_path}")
            
        except Exception as e:
            print(f"Error changing icon: {e}")
            traceback.print_exc()
            self._show_popup(f"Error changing icon: {e}")
            
    def _reload_tray_icon(self):
        """Reloads the tray icon to apply changes immediately."""
        if not self.running:
            return
            
        try:
            # Check if we have an icon
            if not hasattr(self, 'icon') or not self.icon:
                print("No icon to reload")
                return
                
            # Remember the visibility state
            was_visible = False
            if hasattr(self.icon, 'visible'):
                was_visible = self.icon.visible
                
            # Stop the current icon (this automatically removes it from tray)
            print("Stopping current tray icon...")
            try:
                self.icon.stop()
                # Give it a moment to close properly
                time.sleep(0.5)
            except Exception as stop_err:
                print(f"Error stopping icon: {stop_err}")
            
            # Generate new icon image
            new_icon_image = self._icon_image()
            if not new_icon_image:
                print("Failed to create new icon image")
                return
                
            # Create and start a new icon
            print("Creating new tray icon...")
            
            # Define menu again (same as in _setup_tray)
            def schedule(f):
                return lambda: self.root.after(0, f) if self.running else None

            is_vis = lambda i: self.running and self.root.winfo_viewable()
            is_hid = lambda i: self.running and not self.root.winfo_viewable()
            can_clr = lambda i: self.running and bool(self.history)
            can_sv = lambda i: self.running and bool(self.filtered_history_indices)
            is_p = lambda i: self.running and self.capture_paused
            is_r = lambda i: self.running and not self.capture_paused

            menu = (
                item(
                    "Show WinCB-Elite",
                    schedule(self._do_show_window),
                    default=True,
                    enabled=is_hid,
                ),
                item("Hide WinCB-Elite", schedule(self._hide_window), enabled=is_vis),
                item("Pause Capture", schedule(self._toggle_capture), enabled=is_r),
                item("Resume Capture", schedule(self._toggle_capture), enabled=is_p),
                item("Auto-Pause Settings", schedule(self._configure_auto_pause)),
                pystray.Menu.SEPARATOR,
                item(
                    "Snap Window",
                    pystray.Menu(
                        item("Top-Left", schedule(self.snap_top_left)),
                        item("Top-Right", schedule(self.snap_top_right)),
                        item("Bottom-Left", schedule(self.snap_bottom_left)),
                        item("Bottom-Right", schedule(self.snap_bottom_right)),
                        item("Center", schedule(self.snap_center)),
                    ),
                    enabled=is_vis,
                ),
                pystray.Menu.SEPARATOR,
                item(
                    "Export Data", schedule(self._prompt_and_save_batch), enabled=can_sv
                ),
                item("Clear History", schedule(self.clear_history), enabled=can_clr),
                pystray.Menu.SEPARATOR,
                item("Quit WinCB-Elite", schedule(self._prompt_close), enabled=self.running),
            )
            
            # Create new icon
            new_icon = pystray.Icon(
                APP_NAME, new_icon_image, f"{APP_NAME} - History", pystray.Menu(*menu)
            )
            
            # Start the icon in a new thread
            icon_thread = threading.Thread(
                target=lambda: self._run_icon_thread(new_icon, was_visible),
                daemon=True
            )
            icon_thread.start()
            
            print("Icon reloaded successfully")
            
        except Exception as e:
            print(f"Error reloading tray icon: {e}")
            traceback.print_exc()
    
    def _run_icon_thread(self, icon, make_visible=False):
        """Runs the icon in a separate thread and handles visibility."""
        try:
            self.icon = icon
            
            # Show the icon if it was visible before
            if make_visible:
                self.icon.visible = True
                
            print("Starting new tray icon...")
            self.icon.run()
        except Exception as e:
            print(f"Error running icon thread: {e}")
            traceback.print_exc()
        finally:
            print("Icon thread finished")
            self.icon = None

    def _update_group_display(self):
        """Updates the group name display based on current_group_name."""
        if not self.running:
            return
            
        try:
            if self.current_group_name:
                self.group_name_var.set(f"Group: {self.current_group_name}")
                self.group_label.configure(text_color="#4a95eb")  # Bright blue
            else:
                self.group_name_var.set("Default Clip File Showing")
                self.group_label.configure(text_color="#888888")  # Gray
        except Exception as e:
            print(f"Error updating group display: {e}")
            
    def _load_tag_colors(self):
        """Load tag color associations from file."""
        tag_colors_path = APP_DATA_DIR / "tag_colors.json"
        if tag_colors_path.exists():
            try:
                with open(tag_colors_path, "r", encoding="utf-8") as f:
                    self.tag_colors = json.load(f)
                print(f"Loaded {len(self.tag_colors)} tag colors from {tag_colors_path}")
            except Exception as e:
                print(f"Error loading tag colors: {e}")
                self.tag_colors = {}
        else:
            # Default tag colors for common tags
            self.tag_colors = {
                "work": "blue",
                "personal": "green",
                "urgent": "red",
                "reference": "purple",
                "project": "yellow"
            }
            self._save_tag_colors()
            
    def _save_tag_colors(self):
        """Save tag color associations to file."""
        tag_colors_path = APP_DATA_DIR / "tag_colors.json"
        try:
            with open(tag_colors_path, "w", encoding="utf-8") as f:
                json.dump(self.tag_colors, f, indent=2)
        except Exception as e:
            print(f"Error saving tag colors: {e}")
            
    def _set_tag_color(self, tag, color):
        """Set the color for a tag."""
        if color in self.TAG_COLORS:
            self.tag_colors[tag] = color
            self._save_tag_colors()
            
    def _get_tag_color(self, tag):
        """Get the color for a tag, defaulting to blue."""
        return self.tag_colors.get(tag, "blue")
        
    def _add_tag_to_current_clip(self, tag):
        """Add a tag to the current clip."""
        if not tag or not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            return
            
        # Update activity time when adding tags
        self._update_activity_time()
            
        original_index = self.filtered_history_indices[self.current_filtered_index]
        if 0 <= original_index < len(self.history):
            # Create tags list if it doesn't exist
            if "tags" not in self.history[original_index]:
                self.history[original_index]["tags"] = []
                
            # Add tag if it's not already there
            if tag not in self.history[original_index]["tags"]:
                self.history[original_index]["tags"].append(tag)
                self.current_clip_tags = self.history[original_index]["tags"]
                self._update_tag_display()
                self._save_history()
                print(f"Added tag '{tag}' to clip")
                
    def _remove_tag_from_current_clip(self, tag):
        """Remove a tag from the current clip."""
        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            return
            
        # Update activity time when removing tags
        self._update_activity_time()
        
        # Define what happens when confirmed
        def do_remove_tag():
            original_index = self.filtered_history_indices[self.current_filtered_index]
            if 0 <= original_index < len(self.history):
                if "tags" in self.history[original_index] and tag in self.history[original_index]["tags"]:
                    self.history[original_index]["tags"].remove(tag)
                    self.current_clip_tags = self.history[original_index]["tags"]
                    self._update_tag_display()
                    self._save_history()
                    print(f"Removed tag '{tag}' from clip")
        
        # Use the centralized confirmation dialog
        self._confirm_action(
            title="Confirm Tag Removal",
            message=f"Are you sure you want to remove the tag '{tag}'?",
            action_callback=do_remove_tag
        )

    def _update_tag_display(self):
        """Updates the tag display area with current clip tags."""
        if not self.running:
            return
            
        try:
            # Clear existing tag buttons
            for child in self.tag_buttons_frame.winfo_children():
                child.destroy()
                
            # Get tags for current clip
            if not hasattr(self, 'current_clip_tags') or self.current_clip_tags is None:
                self.current_clip_tags = []
                
            if 0 <= self.current_filtered_index < len(self.filtered_history_indices):
                original_index = self.filtered_history_indices[self.current_filtered_index]
                if 0 <= original_index < len(self.history):
                    clip = self.history[original_index]
                    self.current_clip_tags = clip.get("tags", [])
                    
            # Create a button for each tag (up to 3 to avoid crowding)
            max_visible_tags = 3
            visible_tags = self.current_clip_tags[:max_visible_tags]
                    
            # Create a button for each tag
            for tag in visible_tags:
                color_name = self._get_tag_color(tag)
                color_value = self.TAG_COLORS.get(color_name, self.TAG_COLORS["blue"])
                
                # Create frame to hold the tag and remove button
                tag_frame = ctk.CTkFrame(self.tag_buttons_frame, fg_color="transparent")
                tag_frame.pack(side="left", padx=(0, 5))
                
                # Make colors more vibrant
                hover_color = self._brighten_color(color_value, 0.2)
                
                # Tag button with appropriate color
                tag_btn = ctk.CTkButton(
                    tag_frame,
                    text=tag,
                    width=0,  # Auto-width
                    height=22,  # Reduced height
                    fg_color=color_value,
                    hover_color=hover_color,
                    corner_radius=12,
                    text_color="#ffffff",
                    font=("Segoe UI", 10),  # Smaller font
                    command=lambda t=tag: self._filter_by_tag(t)
                )
                tag_btn.pack(side="left")
                self._add_tooltip(tag_btn, f"Click to filter by '{tag}' tag")
                
                # Remove tag button (x)
                remove_btn = ctk.CTkButton(
                    tag_frame,
                    text="×",
                    width=22,  # Reduced width
                    height=22,  # Reduced height
                    fg_color=color_value,
                    hover_color=hover_color,
                    corner_radius=12,
                    font=("Segoe UI", 10),  # Smaller font
                    command=lambda t=tag: self._remove_tag_from_current_clip(t)
                )
                remove_btn.pack(side="left", padx=(1, 0))
                
            # If there are more tags than we're showing, add an indicator
            if len(self.current_clip_tags) > max_visible_tags:
                more_count = len(self.current_clip_tags) - max_visible_tags
                more_label = ctk.CTkLabel(
                    self.tag_buttons_frame,
                    text=f"+{more_count} more",
                    text_color="#888888",
                    font=("Segoe UI", 11, "italic")
                )
                more_label.pack(side="left", padx=(5, 0))
                self._add_tooltip(more_label, "There are more tags on this clip")
                
        except Exception as e:
            print(f"Error updating tag display: {e}")
            traceback.print_exc()

    # Add this helper function for brightening colors
    def _brighten_color(self, hex_color, factor=0.2):
        """Brighten a hex color by the given factor."""
        try:
            # Convert hex to RGB
            hex_color = hex_color.lstrip('#')
            rgb = tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))
            
            # Brighten
            brightened = tuple(min(255, int(c + (255 - c) * factor)) for c in rgb)
            
            # Convert back to hex
            return f"#{brightened[0]:02x}{brightened[1]:02x}{brightened[2]:02x}"
        except Exception as e:
            print(f"Error brightening color {hex_color}: {e}")
            return hex_color

    def _filter_by_tag(self, tag):
        """Filters history to show only clips with the specified tag."""
        if not tag:
            return
            
        # Update activity time when filtering by tag
        self._update_activity_time()
        
        print(f"DEBUG: Filtering by tag: '{tag}'")
        
        
        # Save current clip index in case we need to restore it
        current_original_index = -1
        if 0 <= self.current_filtered_index < len(self.filtered_history_indices):
            try:
                current_original_index = self.filtered_history_indices[self.current_filtered_index]
                print(f"DEBUG: Current clip original index: {current_original_index}")
            except Exception as e:
                print(f"ERROR getting current clip index: {e}")
        
        # Set flags to indicate we're doing tag filtering
        # This prevents tag removal when search is cleared
        self._is_tag_filtering = True
        self._current_filter_tag = tag
            
        # Update the search field to show the tag filter
        self.search_var.set(f"#{tag}")  # Use # prefix to indicate tag search
        
        # Filter history to only include clips with this tag
        self.filtered_history_indices = []
        for i, item in enumerate(self.history):
            if "tags" in item and tag in item.get("tags", []):
                self.filtered_history_indices.append(i)
        
        # Count how many clips have this tag        
        tag_count = len(self.filtered_history_indices)
        print(f"DEBUG: Found {tag_count} clips with tag '{tag}'")
                
        # Update the display to show filtered results
        if tag_count > 0:
            # If the current clip has this tag, try to keep it selected
            new_index = 0
            if current_original_index >= 0:
                try:
                    if current_original_index in self.filtered_history_indices:
                        new_index = self.filtered_history_indices.index(current_original_index)
                        print(f"DEBUG: Current clip has this tag, keeping it selected at position {new_index}")
                except Exception:
                    pass
                    
            self.current_filtered_index = new_index
            self._show_clip()
            self._show_popup(f"Showing {tag_count} clips with tag '{tag}'")
        else:
            self._show_popup(f"No clips found with tag '{tag}'")
            # Don't change the filter - stay in the current view
            # This prevents accidental data loss when clicking on a tag that doesn't exist
            print("DEBUG: No clips with this tag - not changing the current view")
            
            # Just in case we got into an inconsistent state, reset to show all
            if not self.filtered_history_indices:
                print("DEBUG: No clips in filtered view - resetting to show all")
                self.filtered_history_indices = list(range(len(self.history)))
                self.current_filtered_index = 0 if self.filtered_history_indices else -1
                self._is_tag_filtering = False
                
                # Also update the search field to show no filter
                self.search_var.set("")
                
                self._show_clip()
    
    def _show_tag_dialog(self):
        """Shows a dialog to add a new tag with color selection."""
        if not self.running:
            return
            
        # Check if a clip is selected
        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            self._show_popup("Please select a clip first.")
            return
            
        # Create dialog
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Add Tag")
        dialog.geometry("462x560")  # Increased size by 40%
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.focus_set()
        
        # Make dialog modal
        dialog.protocol("WM_DELETE_WINDOW", dialog.destroy)
        
        # Tag entry with better labeling
        entry_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        entry_frame.pack(pady=(15, 15), padx=15, fill="x")
        
        # Bold label for "New Tag Name:"
        new_tag_label = ctk.CTkLabel(
            entry_frame, 
            text="New Tag Name:", 
            font=("Segoe UI", 12, "bold"),
            width=120,
            anchor="w"
        )
        new_tag_label.pack(side="top", anchor="w", pady=(0, 5))
        
        # Tag name entry below the label, full width
        tag_var = ctk.StringVar()
        tag_entry = ctk.CTkEntry(
            entry_frame, 
            textvariable=tag_var, 
            width=432,  # Full width of dialog minus padding
            height=30   # Taller for better visibility
        )
        tag_entry.pack(fill="x", pady=(0, 0))
        
        # Get all existing tags
        existing_tags = set()
        for item in self.history:
            if "tags" in item:
                existing_tags.update(item.get("tags", []))
        
        if existing_tags:
            # Add a search box for filtering tags using same structure as name field
            filter_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            filter_frame.pack(pady=(0, 15), padx=15, fill="x")
            
            # Match the label structure of the New Tag Name field
            filter_label = ctk.CTkLabel(
                filter_frame, 
                text="Filter Tags:", 
                font=("Segoe UI", 12, "bold"),
                width=120,
                anchor="w"
            )
            filter_label.pack(side="top", anchor="w", pady=(0, 5))
            
            # Filter entry with same structure as name entry
            filter_var = ctk.StringVar()
            filter_entry = ctk.CTkEntry(
                filter_frame,
                textvariable=filter_var, 
                width=432,
                height=30,
                placeholder_text="Type to filter tags..."
            )
            filter_entry.pack(fill="x", pady=(0, 0))
            
            # Existing tags section
            existing_tags_frame = ctk.CTkFrame(dialog, fg_color="transparent")
            existing_tags_frame.pack(pady=(0, 15), padx=15, fill="x")
            
            ctk.CTkLabel(
                existing_tags_frame,
                text="Existing tags (click to use):",
                font=("Segoe UI", 12, "bold"),
                anchor="w"
            ).pack(side="top", anchor="w", pady=(0, 10))
            
            # Create a frame for the tag buttons
            tags_buttons_frame = ctk.CTkFrame(existing_tags_frame, fg_color="transparent")
            tags_buttons_frame.pack(fill="x")
            
            # List of existing tags for the user to click
            row_tags = 0
            col_tags = 0
            max_cols = 2
            
            sorted_tags = sorted(list(existing_tags))
            for tag in sorted_tags:
                color_name = self._get_tag_color(tag)
                color_value = self.TAG_COLORS.get(color_name, self.TAG_COLORS["blue"])
                hover_color = self._brighten_color(color_value, 0.2)
                
                def set_tag_name(t=tag, c=color_name):
                    tag_var.set(t)
                    color_var.set(c)  # Also set the color to match the tag's current color
                
                tag_btn = ctk.CTkButton(
                    tags_buttons_frame,
                    text=tag,
                    fg_color=color_value,
                    hover_color=hover_color,
                    text_color="#ffffff",
                    height=28,
                    corner_radius=14,
                    command=set_tag_name
                )
                tag_btn.grid(row=row_tags, column=col_tags, padx=5, pady=3, sticky="ew")
                
                col_tags += 1
                if col_tags >= max_cols:
                    col_tags = 0
                    row_tags += 1
            
            # Configure grid columns to be equal width
            tags_buttons_frame.grid_columnconfigure(0, weight=1)
            tags_buttons_frame.grid_columnconfigure(1, weight=1)
        
        # Tag colors section
        color_section = ctk.CTkFrame(dialog, fg_color="transparent")
        color_section.pack(pady=(5, 15), padx=15, fill="x")
        
        ctk.CTkLabel(
            color_section, 
            text="Tag color:", 
            font=("Segoe UI", 12, "bold"),
            anchor="w"
        ).pack(pady=(0, 10), anchor="w")
        
        # Color selection frame
        colors_frame = ctk.CTkFrame(color_section, fg_color="transparent")
        colors_frame.pack(pady=(0, 10), fill="x")
        
        # Color selection
        color_var = ctk.StringVar(value="blue")  # Default color
        
        # Create a button for each color
        for i, (color_name, color_value) in enumerate(self.TAG_COLORS.items()):
            color_btn = ctk.CTkRadioButton(
                colors_frame,
                text=color_name.capitalize(),
                variable=color_var,
                value=color_name,
                fg_color=color_value,
                hover_color=color_value,
                border_color=color_value
            )
            color_btn.grid(row=i//3, column=i%3, padx=8, pady=5, sticky="w")
        
        # Add button to add the tag
        def add_tag():
            tag = tag_var.get().strip()
            color = color_var.get()
            
            if not tag:
                return
                
            # Set tag color mapping
            self._set_tag_color(tag, color)
            
            # Add tag to current clip
            self._add_tag_to_current_clip(tag)
            
            # Close dialog
            dialog.destroy()
            
        # Buttons
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(pady=(15, 20), fill="x", padx=15)
        
        ctk.CTkButton(
            buttons_frame, text="Cancel", width=120, height=36,
            command=dialog.destroy
        ).pack(side="left", padx=(0, 10))
        
        ctk.CTkButton(
            buttons_frame, text="Add Tag", width=120, height=36,
            command=add_tag
        ).pack(side="right")
        
        # Center dialog
        self._center_toplevel(dialog)
        
        # Set focus to entry
        tag_entry.focus_set()
            
    def _clear_current_clip_tags(self):
        """Clears all tags from the current clip."""
        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            return
            
        # Show confirmation dialog before clearing tags
        self._confirm_action(
            "Delete All Tags", 
            "Are you sure you want to delete all known clip tags. Continue or Cancel?",
            lambda: self._execute_clear_tags()
        )
        
    def _execute_clear_tags(self):
        """Actually performs the tag clearing operation after confirmation."""
        if not (0 <= self.current_filtered_index < len(self.filtered_history_indices)):
            return
            
        original_index = self.filtered_history_indices[self.current_filtered_index]
        if 0 <= original_index < len(self.history):
            if "tags" in self.history[original_index]:
                self.history[original_index]["tags"] = []
                self.current_clip_tags = []
                self._update_tag_display()
                self._save_history()
                print("All tags cleared from current clip")
                self._show_popup("All tags deleted from this clip")

    def _configure_auto_pause(self):
        """Opens a dialog to configure the auto-pause timer."""
        if not self.running:
            return
            
        # Always read fresh value from config first
        try:
            self._load_config()  # Reload config to get latest value
            print(f"Auto-pause dialog opened with setting: {self.auto_pause_seconds} seconds")
        except Exception as e:
            print(f"Error loading config in auto-pause dialog: {e}")
            
        # Create dialog window
        dialog = ctk.CTkToplevel(self.root)
        dialog.title("Auto-Pause Settings")
        dialog.geometry("350x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)
        
        # Heading
        ctk.CTkLabel(
            dialog,
            text="Auto-Pause Configuration",
            font=("Segoe UI", 14, "bold")
        ).pack(pady=(15, 5))
        
        # Explanation text - updated to clarify behavior
        ctk.CTkLabel(
            dialog,
            text="Automatically pause clipboard capture after the specified period of inactivity.\nSet to 0 to disable auto-pause.",
            font=("Segoe UI", 11),
            wraplength=300
        ).pack(pady=(0, 10))
        
        # Input frame
        input_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        input_frame.pack(pady=(0, 15), fill="x", padx=20)
        
        # Label
        ctk.CTkLabel(
            input_frame,
            text="Auto-Pause in (seconds):",
            font=("Segoe UI", 12),
            width=170,
            anchor="e"
        ).pack(side="left", padx=(0, 10))
        
        # Entry for seconds
        seconds_var = ctk.StringVar(value=str(self.auto_pause_seconds))
        seconds_entry = ctk.CTkEntry(
            input_frame,
            textvariable=seconds_var,
            width=80,
            justify="center"
        )
        seconds_entry.pack(side="left")
        
        # Focus the entry
        seconds_entry.focus_set()
        
        # Buttons frame
        buttons_frame = ctk.CTkFrame(dialog, fg_color="transparent")
        buttons_frame.pack(pady=(5, 15), fill="x", padx=20)
        
        # Cancel button
        ctk.CTkButton(
            buttons_frame,
            text="Cancel",
            width=100,
            command=dialog.destroy
        ).pack(side="left", padx=(0, 10))
        
        # Set button - applies the new setting
        def on_set():
            try:
                seconds = int(seconds_var.get())
                if seconds < 0:
                    seconds = 0
                
                self.auto_pause_seconds = seconds
                
                # Cancel existing timer if any
                if self.auto_pause_timer:
                    self.root.after_cancel(self.auto_pause_timer)
                    self.auto_pause_timer = None
                
                # Start new timer if enabled and clipboard monitoring is active
                if seconds > 0 and not self.capture_paused:
                    self._start_auto_pause_timer()
                    
                # Save setting to config (always update and save immediately)
                self.config["auto_pause_seconds"] = seconds
                print(f"Saving auto-pause setting: {seconds} seconds")
                self._save_config()  # Save immediately

                # Make double-sure it's saved to config file
                try:
                    # Create a separate direct update to config file as backup
                    config_file_path = CONFIG_FILE_PATH
                    if os.path.exists(config_file_path):
                        with open(config_file_path, "r", encoding="utf-8") as f:
                            config_data = json.load(f)
                        config_data["auto_pause_seconds"] = seconds
                        with open(config_file_path, "w", encoding="utf-8") as f:
                            json.dump(config_data, f, ensure_ascii=False, indent=2)
                        print(f"Auto-pause setting written directly to config file: {seconds} seconds")
                except Exception as direct_save_err:
                    print(f"Error during direct config file write: {direct_save_err}")
                
                # Show feedback and close dialog
                dialog.destroy()
                
                if seconds > 0:
                    self._show_popup(f"Auto-pause set to {seconds} seconds")
                else:
                    self._show_popup("Auto-pause disabled")
                
            except ValueError:
                self._show_popup("Please enter a valid number")
        
        ctk.CTkButton(
            buttons_frame,
            text="Set",
            width=100,
            command=on_set
        ).pack(side="right")
        
        # Center the dialog
        self._center_toplevel(dialog)
    
    def _update_activity_time(self):
        """Updates the last activity timestamp to indicate user is active in the app.
        This should be called for any significant user interaction."""
        self.last_activity_time = time.time()
        
    def _start_auto_pause_timer(self):
        """Starts the auto-pause timer that checks for user inactivity."""
        if not self.running or self.capture_paused or self.auto_pause_seconds <= 0:
            return
        
        # Cancel existing timer if any
        if self.auto_pause_timer:
            self.root.after_cancel(self.auto_pause_timer)
        
        # Define the auto-pause check function
        def check_inactivity():
            if not self.running or self.capture_paused:
                return
            
            current_time = time.time()
            idle_time = current_time - self.last_activity_time
            
            if idle_time >= self.auto_pause_seconds:
                # No activity for the specified duration, auto-pause now
                self.capture_paused = True
                self.pause_resume_btn.configure(
                    text="Resume Capture", 
                    fg_color="orange", 
                    hover_color="darkorange"
                )
                print(f"Clipboard capture AUTO-PAUSED after {int(idle_time)} seconds of inactivity. Use 'Resume Capture' button or tray menu to resume.")
                self._show_popup(f"Clipboard capture auto-paused after {int(idle_time)} seconds of inactivity")
                self.auto_pause_timer = None
            else:
                # Still active, continue checking
                remaining = self.auto_pause_seconds - idle_time
                print(f"DEBUG: Auto-pause: {int(remaining)} seconds of inactivity remaining before auto-pause")
                # Schedule next check (every 10 seconds for efficiency)
                self.auto_pause_timer = self.root.after(10000, check_inactivity)
        
        # Start the inactivity check (first check after 10 seconds)
        self.auto_pause_timer = self.root.after(10000, check_inactivity)
        print(f"Auto-pause timer started: Will pause after {self.auto_pause_seconds} seconds of inactivity")


# --- Main Execution ---
if __name__ == "__main__":
    print(f"--- Starting {APP_NAME} --- {time.strftime('%Y-%m-d %H:%M:%S')}")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        print("DPI Awareness set (1)")
    except Exception as dpi_e:
        print(f"Info: DPI failed - {dpi_e}")
    try:
        app = WinCB_Elite()
    except KeyboardInterrupt:
        print("\n--- Kbd Interrupt, exiting ---")
        sys.exit(0)
    except Exception as e:
        print(f"\n--- FATAL ERROR ---")
        print(f"{e}")
        traceback.print_exc()
        try:
            err = ctk.CTk()
            err.withdraw()
            from tkinter import messagebox

            messagebox.showerror(
                f"{APP_NAME} Fatal Error", f"Error:\n\n{e}\n\nSee console."
            )
            err.destroy()
        except:
            pass
        sys.exit(1)
    finally:
        print(f"--- {APP_NAME} finished --- {time.strftime('%Y-%m-d %H:%M:%S')}")
