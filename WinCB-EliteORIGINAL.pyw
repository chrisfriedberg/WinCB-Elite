# Clippy_part1.py
# Part 1 of 5 for Clippy.pyw
# Contains imports, configuration, and Clippy class initialization up to _setup_sidebar_buttons
# To combine: Concatenate part1 + part2 + part3 + part4 + part5 into Clippy.pyw

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
APP_NAME = "Clippy (Nitro)"
HISTORY_LIMIT = 50
AUTOSAVE_DELAY_MS = 1500  # Delay for auto-saving text edits (1.5 seconds)
MAIN_BG_COLOR = "#2b2b2b"  # Main dark background color
TEXT_FG_COLOR = "#ffffff"  # White text color
CURSOR_COLOR = "#ffffff"  # White cursor color
HIGHLIGHT_COLOR = "cyan"  # Highlight border for focused text widget

# --- Path Configuration ---
# Use user's home directory for persistent data
try:
    APP_DATA_DIR = Path.home() / "Clippy"
    APP_DATA_DIR.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE_PATH = APP_DATA_DIR / "clippy_history.json"
    BATCH_SAVE_DIR = APP_DATA_DIR
    print(f"Using data directory: {APP_DATA_DIR}")
except Exception as path_e:
    print(f"FATAL: Could not create or access data directory: {Path.home() / 'Clippy'}")
    print(f"Error: {path_e}")
    # Fallback to script directory if home directory fails (less ideal for EXEs)
    try:
        script_dir = Path(os.path.dirname(os.path.abspath(__file__)))
    except NameError:
        script_dir = Path(os.getcwd())
    APP_DATA_DIR = script_dir
    HISTORY_FILE_PATH = APP_DATA_DIR / "clippy_history.json"
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


class Clippy:
    """A clipboard manager application with history, text/image support, and system tray integration."""

    def __init__(self):
        """Initialize the Clippy application window and components."""
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("dark-blue")

        self.root = ctk.CTk()
        self.root.title(APP_NAME)
        self.root.geometry("900x700+650+0")
        self.root.minsize(750, 550)
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

        # Simplify search bar: remove placeholder
        self.search_var = ctk.StringVar()
        self.search_entry = ctk.CTkEntry(
            self.top_nav, textvariable=self.search_var, width=120
        )
        self.search_entry.grid(row=0, column=0, padx=(0, 5), pady=5, sticky="w")
        self.search_entry.bind("<KeyRelease>", self._on_search_change)
        self._add_tooltip(self.search_entry, "Search clip titles and text content")

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
        self._load_history()
        self._filter_and_show()
        self.root.after(500, self.poll_clipboard)
        self.icon_thread = threading.Thread(target=self._setup_tray, daemon=True)
        self.icon_thread.start()

        print("Clippy UI Initialized.")
        self.root.mainloop()

    # --- Setup Helpers ---

    def _setup_sidebar_buttons(self):
        """Creates and packs all buttons in the sidebar frame."""
        self._add_sidebar_separator("Navigation")
        # Frame to hold the Start and End buttons side by side
        nav_jump_frame = ctk.CTkFrame(self.sidebar_frame, fg_color="transparent")
        nav_jump_frame.pack(pady=5, padx=10, anchor="n")

        # End button (jump to oldest clip, last index) - Now on the left
        self.end_btn = ctk.CTkButton(
            nav_jump_frame, text="|<", width=55, height=30, command=self.jump_to_oldest
        )
        self.end_btn.pack(side="left", padx=(0, 5))
        self._add_tooltip(self.end_btn, "Jump to oldest clip (Alt+End)")

        # Start button (jump to newest clip, index 0) - Now on the right
        self.start_btn = ctk.CTkButton(
            nav_jump_frame, text=">|", width=55, height=30, command=self.jump_to_newest
        )
        self.start_btn.pack(side="right")
        self._add_tooltip(self.start_btn, "Jump to newest clip (Alt+Home)")

        # Older and Newer buttons
        self.older_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="\u2190 Older",
            width=120,
            height=30,
            command=self.prev_clip,
        )
        self.older_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.older_btn, "View previous clip in history (Alt+Left)")
        self.newer_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Newer \u2192",
            width=120,
            height=30,
            command=self.next_clip,
        )
        self.newer_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.newer_btn, "View next clip in history (Alt+Right)")
        self.delete_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Delete Clip",
            width=120,
            height=30,
            command=self.delete_current_clip,
        )
        self.delete_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.delete_btn, "Delete the currently displayed clip (Delete)"
        )

        self._add_sidebar_separator("Management")
        self.copy_sel_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Copy Selection",
            width=120,
            height=30,
            command=self.copy_selection_to_history,
        )
        self.copy_sel_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.copy_sel_btn, "Copy selected text from the editor as a new clip"
        )
        self.copy_clip_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Copy Clip",
            width=120,
            height=30,
            command=self.copy_clip_to_clipboard,
        )
        self.copy_clip_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.copy_clip_btn,
            "Copy the entire current clip back to the system clipboard (Ctrl+C)",
        )
        self.titles_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Preview Titles",
            width=120,
            height=30,
            command=self._show_titles_modal,
        )
        self.titles_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.titles_btn, "View/select clips by title (with previews)")
        self.clear_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Clear History",
            width=120,
            height=30,
            command=self.clear_history,
        )
        self.clear_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.clear_btn, "Delete all clips from history (confirmation required)"
        )
        self.hide_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Hide",
            width=120,
            height=30,
            command=self._hide_window,
        )
        self.hide_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(self.hide_btn, "Minimize Clippy to the system tray")

        self._add_sidebar_separator("File Management")
        # Set width to match the longest button text ("Save/Restore Clip Group")
        bottom_button_width = 160  # Matches "Save/Restore Clip Group"
        self.save_batch_btn = ctk.CTkButton(
            self.sidebar_frame,
            text="Export Data",
            width=bottom_button_width,
            height=30,
            command=self._prompt_and_save_batch,
        )
        self.save_batch_btn.pack(pady=5, padx=10, anchor="n")
        self._add_tooltip(
            self.save_batch_btn,
            f"Export currently filtered clips to a text file in\n{BATCH_SAVE_DIR}",
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
        self._add_tooltip(self.close_btn, "Close Clippy (confirmation required)")

        # --- Key Bindings ---
        self.root.bind("<Alt-Left>", lambda e: self.prev_clip())
        self.root.bind("<Alt-Right>", lambda e: self.next_clip())
        # Add bindings for Start (newest) and End (oldest)
        self.root.bind("<Alt-Home>", lambda e: self.jump_to_newest())
        self.root.bind("<Alt-End>", lambda e: self.jump_to_oldest())

        # Conditional Delete key binding to avoid deleting clips while editing
        def delete_clip_if_no_focus(event):
            focused = self.root.focus_get()
            if focused not in (self.title_entry, self.textbox, self.search_entry):
                self.delete_current_clip()

        self.root.bind("<Delete>", delete_clip_if_no_focus)
        self.root.bind("<Control-c>", lambda e: self.copy_clip_to_clipboard())
        self.root.bind("<Control-C>", lambda e: self.copy_clip_to_clipboard())
        self.root.bind("<Control-f>", lambda e: self.search_entry.focus_set())
        self.root.bind("<Control-F>", lambda e: self.search_entry.focus_set())

    def _on_search_change(self, event=None):
        """Handle search input changes."""
        self._finalize_text_edit()
        self._filter_and_show()

    # Clippy_part2.py
    # Part 2 of 5 for Clippy.pyw
    # Contains sidebar helpers and history persistence methods
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into Clippy.pyw

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

    # Clippy_part3.py
    # Part 3 of 5 for Clippy.pyw
    # Contains clipboard polling and filtering methods
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into Clippy.pyw

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
            print("Clipboard capture PAUSED.")
        else:
            self.pause_resume_btn.configure(
                text="Pause Capture", fg_color=default_color, hover_color=hover_color
            )
            print("Clipboard capture RESUMED.")

    def poll_clipboard(self):
        """Poll system clipboard for new content."""
        if not self.running:
            return

        now = time.time()

        if self.capture_paused:
            if self.running:
                self.root.after(1000, self.poll_clipboard)
            return

        if now < self.ignore_clip_until:
            if self.running:
                self.root.after(200, self.poll_clipboard)
            return

        new_clip_found = False
        clip_type = None
        clip_data = None

        try:
            win32clipboard.OpenClipboard()

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
                        print(f"Warning: CF_UNICODETEXT failed ({te}), trying CF_TEXT.")
                        if has_text:
                            try:
                                text_data = win32clipboard.GetClipboardData(
                                    win32clipboard.CF_TEXT
                                ).decode("mbcs")
                            except Exception as decode_err:
                                print(f"Warning: CF_TEXT decode failed: {decode_err}")
                        else:
                            text_data = None
                    except Exception as e:
                        print(f"Error getting CF_UNICODETEXT: {e}")
                        text_data = None
                elif has_text:
                    try:
                        text_data = win32clipboard.GetClipboardData(
                            win32clipboard.CF_TEXT
                        ).decode("mbcs")
                    except Exception as e:
                        print(f"Error getting/decoding CF_TEXT: {e}")

                if (
                    text_data is not None
                    and text_data.strip()
                    and text_data != self.last_clip_text
                ):
                    if len(text_data) > 500_000:
                        print(
                            f"Skipped excessively long text clip ({len(text_data)} chars)"
                        )
                    else:
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
                            print(f"New image data detected: {len(image_data)} bytes")
                            self.last_clip_img_data = image_data
                            self.last_clip_text = (
                                None  # Important to clear the other type
                            )
                            new_clip_found = True
                            clip_type = "image"
                            clip_data = image_data
                        # else:
                        # You probably don't need the "unchanged" message printed repeatedly.
                        # If you want it for debugging, uncomment the next line.
                        # print("Image data unchanged.")
                        # pass # No action needed if data is the same
                except Exception as e:
                    print(f"Error getting CF_DIB data: {e}")
                    traceback.print_exc()

            win32clipboard.CloseClipboard()

            if new_clip_found and clip_type and clip_data is not None:
                log_preview = ""
                if clip_type == "text":
                    log_preview = clip_data[:60].replace("\n", "\\n").replace(
                        "\r", ""
                    ) + ("..." if len(clip_data) > 60 else "")
                    print(f"Captured text: {log_preview}")
                elif clip_type == "image":
                    print(f"Captured image ({len(clip_data)} bytes)")
                action = lambda ct=clip_type, cd=clip_data: self._add_to_history(ct, cd)
                self.root.after(0, action)

        except win32clipboard.error as e:
            if e.winerror != 5:
                print(f"Clipboard access error (winerror {e.winerror}): {e}")
            try:
                win32clipboard.CloseClipboard()
            except:
                pass
        except Exception as e:
            print(f"Error during clipboard polling: {e}")
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
        if not self.running or data is None:
            print(
                f"Failed to add to history: running={self.running}, data={data is None}"
            )
            return

        current_time = time.time()
        new_entry = {
            "type": data_type,
            "content": data,
            "title": "",
            "timestamp": current_time,
        }

        is_duplicate = False
        if not is_from_selection and self.history:
            last_entry = self.history[0]
            last_ts = last_entry.get("timestamp", 0)
            if current_time - last_ts < 0.8:
                is_duplicate = True
                print(
                    f"Skipped: Too soon after last capture ({current_time - last_ts:.2f}s)"
                )
            elif data_type == last_entry.get("type"):
                if (
                    data_type == "text"
                    and isinstance(data, str)
                    and isinstance(last_entry.get("content"), str)
                ):
                    if data.strip() == last_entry["content"].strip():
                        is_duplicate = True
                        print("Skipped: Duplicate text content")
                elif (
                    data_type == "image"
                    and isinstance(data, bytes)
                    and data == last_entry.get("content")
                ):
                    is_duplicate = True
                    print("Skipped: Duplicate image content")

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
                print(f"Image title generated: {new_entry['title']}")
            except Exception as e:
                print(f"Failed to process image for title: {e}")
                traceback.print_exc()
                new_entry["title"] = f"Image_{time.strftime('%Y%m%d_%H%M%S')}_[ERR]"
        elif data_type == "text" and isinstance(data, str):
            stripped_data = data.strip()
            if not stripped_data:
                print("Skipped empty text entry.")
                return
            first_line = stripped_data.split("\n", 1)[0]
            max_title_len = 60  # Increased to capture more context
            title = re.sub(r"\s+", " ", first_line).strip()
            new_entry["title"] = (
                (title[:max_title_len] + "...") if len(title) > max_title_len else title
            )
            # Remove trailing periods or ellipses for clarity
            new_entry["title"] = new_entry["title"].rstrip(".").rstrip("...")
        else:
            print(
                f"Warning: Invalid data type '{data_type}' or data provided to _add_to_history."
            )
            return

        if len(self.history) >= HISTORY_LIMIT:
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
                elif (
                    new_entry["type"] == "text"
                    and search_term in new_entry["content"].lower()
                ):
                    matches_search = True
            except Exception as search_check_err:
                print(
                    f"Warning: Error checking if new item matches search: {search_check_err}"
                )

            if not matches_search:
                should_update_display_fully = False
                self._filter_history()
                self._update_page_label()
                print(
                    "New item added but doesn't match current filter. Display not changed."
                )

        if should_update_display_fully:
            self._filter_history()
            if self.filtered_history_indices and self.filtered_history_indices[0] == 0:
                self.current_filtered_index = 0
                self._show_clip()
            else:
                self._filter_and_show()

    # --- Filtering and Display ---

    def _on_search_change(self, event=None):
        """Handle search input changes."""
        self._finalize_text_edit()
        self._filter_and_show()

    def _filter_history(self):
        """Filter history based on search query."""
        search_term = self.search_entry.get().lower().strip()
        if not search_term:
            self.filtered_history_indices = list(range(len(self.history)))
        else:
            self.filtered_history_indices = []
            for i, item in enumerate(self.history):
                item_matches = False
                try:
                    if search_term in item.get("title", "").lower():
                        item_matches = True
                    elif item.get("type") == "text" and isinstance(
                        item.get("content"), str
                    ):
                        if search_term in item["content"].lower():
                            item_matches = True
                    if item_matches:
                        self.filtered_history_indices.append(i)
                except Exception as e:
                    print(f"Warning: Error during search filter on item index {i}: {e}")

    def _filter_and_show(self):
        """Filter history and update display, trying to preserve current item if possible."""
        current_original_index = -1
        if 0 <= self.current_filtered_index < len(self.filtered_history_indices):
            try:
                current_original_index = self.filtered_history_indices[
                    self.current_filtered_index
                ]
            except IndexError:
                pass

        self._filter_history()

        num_filtered = len(self.filtered_history_indices)

        if num_filtered == 0:
            self.current_filtered_index = -1
        else:
            new_filtered_index = -1
            if current_original_index != -1:
                try:
                    new_filtered_index = self.filtered_history_indices.index(
                        current_original_index
                    )
                except ValueError:
                    pass
            self.current_filtered_index = (
                new_filtered_index if new_filtered_index != -1 else 0
            )

        self._show_clip()

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

    # Clippy_part4.py
    # Part 4 of 5 for Clippy.pyw
    # Contains display, text editing, button actions, and titles modal methods
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into Clippy.pyw

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
                        self.textbox.configure(state="disabled")
                    except Exception as img_e:
                        print(f"Error setting image label: {img_e}")
                        display_message = "Error displaying image preview."
                        self.current_display_image = None
                        if self.img_label.winfo_ismapped():
                            self.img_label.pack_forget()
                else:
                    display_message = "Image preview unavailable or invalid."
                    self.textbox.configure(state="disabled")

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
        self.root.update_idletasks()
        self._update_scrollregion()

    # --- Text Editing ---

    def _on_text_edited(self, event=None):
        """Handles text modification in the textbox, schedules auto-save."""
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

            if self.history[original_index]["type"] != "text":
                print("Warning: Attempted to save text edit to non-text item.")
                return

            new_content = self.textbox.get("1.0", "end-1c")

            if self.history[original_index]["content"] != new_content:
                print(
                    f"Auto-saving changes to clip: {self.history[original_index]['title']}"
                )
                self.history[original_index]["content"] = new_content
                self.history[original_index]["timestamp"] = time.time()
                self._save_history()
                self.textbox.edit_modified(False)
                self.current_clip_modified = False
            else:
                self.textbox.edit_modified(False)
                self.current_clip_modified = False

        except IndexError:
            print("Error saving text: Index out of bounds (race condition?).")
        except Exception as e:
            print(f"Error saving edited text: {e}")
            traceback.print_exc()

    def _finalize_text_edit(self):
        """Immediately saves any pending text edits if the timer is active."""
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
        num_filtered = len(self.filtered_history_indices)
        if num_filtered <= 1:
            self._show_popup("No older clips available.")
            return
        self.current_filtered_index = num_filtered - 1  # Last index is the oldest
        self._show_clip()

    def jump_to_newest(self):
        """Moves to the newest clip in the filtered history."""
        self._finalize_text_edit()
        num_filtered = len(self.filtered_history_indices)
        if num_filtered <= 1:
            self._show_popup("No newer clips available.")
            return
        self.current_filtered_index = 0  # Index 0 is the newest
        self._show_clip()

    def delete_current_clip(self):
        """Deletes the currently displayed clip from history."""
        self._finalize_text_edit()

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

            removed_item = self.history.pop(original_index_to_delete)
            print(f"Deleted clip: {removed_item.get('title', '?')}")
            self._save_history()
            self._filter_and_show()

        except IndexError:
            self._show_popup("Error deleting clip: Invalid index.")
        except Exception as e:
            print(f"Error deleting clip: {e}")
            traceback.print_exc()
            self._filter_and_show()

    def copy_selection_to_history(self):
        """Copies selected text from the textbox as a new history item."""
        if (
            not self.running
            or not self.textbox.winfo_ismapped()
            or self.textbox.cget("state") == "disabled"
        ):
            self._show_popup("No editable text selected.")
            return
        try:
            selection = self.textbox.get("sel.first", "sel.last")
        except TclError:
            selection = ""
        if selection and selection.strip():
            self._add_to_history("text", selection, is_from_selection=True)
            self._force_copy_to_clipboard("text", selection)
            self._show_popup("Selection added as new clip and copied to clipboard.")
        else:
            self._show_popup("No text selected to copy.")

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
            content = item.get("content")

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
        """Helper to forcefully set the system clipboard."""
        try:
            win32clipboard.OpenClipboard()
            win32clipboard.EmptyClipboard()
            if kind == "text" and isinstance(content, str):
                win32clipboard.SetClipboardData(win32clipboard.CF_UNICODETEXT, content)
                self.last_clip_text = content
                self.last_clip_img_data = None
            elif kind == "image" and isinstance(content, bytes):
                win32clipboard.SetClipboardData(CF_DIB, content)
                self.last_clip_img_data = content
                self.last_clip_text = None
            else:
                print(f"Cannot force copy to clipboard: Invalid type {kind}")
                win32clipboard.CloseClipboard()
                return
            win32clipboard.CloseClipboard()
            self.ignore_clip_until = time.time() + 1.5
        except Exception as e:
            print(f"Error forcing clipboard copy: {e}")
            traceback.print_exc()
            self._show_popup(f"Clipboard Error: {e}")
            try:
                win32clipboard.CloseClipboard()
            except:
                pass

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
        dialog.geometry("350x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)
        ctk.CTkLabel(
            dialog,
            text="Clear entire clipboard history?\nThis action cannot be undone.",
        ).pack(pady=20, padx=10)
        btns = ctk.CTkFrame(dialog, fg_color="transparent")
        btns.pack(pady=10)

        def on_yes():
            try:
                print("Clearing clipboard history...")
                self.history = []
                self.filtered_history_indices = []
                self.current_filtered_index = -1
                self.tk_image_references.clear()
                self.search_var.set("")
                self.title_var.set("")
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
            btn.bind(
                "<Enter>",
                lambda ev, i=idx, w=btn: self._show_preview_popup(ev, i, w),
                add="+",
            )
            btn.bind("<Leave>", lambda ev: self._hide_preview_popup(), add="+")
            btn.bind("<Button-1>", lambda ev: self._hide_preview_popup(), add="+")
        self._center_toplevel(modal)

    def _show_preview_popup(self, event, history_index, widget):
        """Creates and shows a Toplevel popup preview."""
        if not self.running or not (0 <= history_index < len(self.history)):
            return
        self._hide_preview_popup()
        try:
            item = self.history[history_index]
            kind, content = item.get("type"), item.get("content")
            if content is None:
                return
            self.preview_popup = Toplevel(self.root)
            self.preview_popup.wm_overrideredirect(True)
            self.preview_popup.attributes("-topmost", True)
            border = Frame(self.preview_popup, bg="#aaa", bd=1)
            border.pack(fill="both", expand=True)
            inner = Frame(border, bg="#333")
            inner.pack(fill="both", expand=True, padx=1, pady=1)
            wid = None
            if kind == "text" and isinstance(content, str):
                lines = content.strip().split("\n")
                txt = "\n".join(lines[:PREVIEW_MAX_TEXT_LINES])
                txt = "\n".join(
                    [
                        l[:PREVIEW_MAX_TEXT_CHARS_PER_LINE]
                        + ("..." if len(l) > PREVIEW_MAX_TEXT_CHARS_PER_LINE else "")
                        for l in txt.split("\n")
                    ]
                )
                if len(lines) > PREVIEW_MAX_TEXT_LINES:
                    txt += "\n..."
                wid = Label(
                    inner,
                    text=txt or "(Empty)",
                    justify="left",
                    bg="#333",
                    fg="#fff" if txt else "#aaa",
                    font=("Segoe UI", 9),
                    padx=5,
                    pady=5,
                    wraplength=PREVIEW_MAX_IMAGE_WIDTH * 1.5,
                )
            elif kind == "image" and isinstance(content, bytes):
                try:
                    img = Image.open(BytesIO(content))
                    img.thumbnail(
                        (PREVIEW_MAX_IMAGE_WIDTH, PREVIEW_MAX_IMAGE_HEIGHT),
                        Image.Resampling.LANCZOS,
                    )
                    tk_img = ImageTk.PhotoImage(img)
                    self.tk_image_references[id(self.preview_popup)] = tk_img
                    wid = Label(inner, image=tk_img, bg="#333")
                except Exception as e:
                    print(f"Warn: Preview img err: {e}")
                    wid = Label(
                        inner,
                        text="(Img Err)",
                        bg="#333",
                        fg="#f88",
                        font=("Segoe UI", 9),
                        padx=5,
                        pady=5,
                    )
            if wid:
                wid.pack()
                self.preview_popup.update_idletasks()
                x = widget.winfo_rootx() + widget.winfo_width() + 10
                y = widget.winfo_rooty()
                self.preview_popup.geometry(f"+{x}+{y}")
            else:
                self._hide_preview_popup()
        except Exception as e:
            print(f"Error preview: {e}")
            traceback.print_exc()
            self._hide_preview_popup()

    def _hide_preview_popup(self):
        """Destroys the preview popup window."""
        if self.preview_popup:
            popup_id = id(self.preview_popup)
            try:
                self.preview_popup.destroy()
            except:
                pass
            if popup_id in self.tk_image_references:
                del self.tk_image_references[popup_id]
        self.preview_popup = None

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

    # Clippy_part5.py
    # Part 5 of 5 for Clippy.pyw
    # Contains batch save, save/restore group, window management, tray setup, and main execution
    # To combine: Concatenate part1 + part2 + part3 + part4 + part5 into Clippy.pyw

    # --- Save Batch ---

    def _sanitize_filename(self, name):
        """Sanitizes a string for use as a filename."""
        name = re.sub(r'[<>:"/\\|?*]', "", name).strip(". ")
        if re.match(r"^(CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])$", name, re.IGNORECASE):
            name = "_" + name
        return name if name else "Untitled_Clippy_Batch"

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
        filename = f"{name}.txt"
        filepath = os.path.join(str(BATCH_SAVE_DIR), filename)

        num_save = len(self.filtered_history_indices)
        texts = 0
        images = []
        try:
            search_text = self.search_entry.get()  # No placeholder check needed
            with open(filepath, "w", encoding="utf-8") as f:
                f.write(
                    f"Clippy Export: {name}\nFilter: '{search_text}'\nClips: {num_save}\n{'='*40}\n\n"
                )
                for filt_idx, orig_idx in enumerate(self.filtered_history_indices):
                    if 0 <= orig_idx < len(self.history):
                        item = self.history[orig_idx]
                        title = item.get("title", "?")
                        kind = item.get("type")
                        content = item.get("content")
                        f.write(
                            f"--- Clip {filt_idx+1}/{num_save} ---\nTitle: {title}\nType: {kind}\n"
                        )
                        if kind == "text" and isinstance(content, str):
                            f.write(f"Content:\n{content}\n{'-'*20}\n\n")
                            texts += 1
                        elif kind == "image":
                            f.write(f"Content: [Image]\n{'-'*20}\n\n")
                            images.append(title)
                        else:
                            f.write(f"Content: [?]\n{'-'*20}\n\n")
                    else:
                        f.write(
                            f"--- Clip {filt_idx+1}/{num_save} ---\nError: Bad Index\n{'-'*20}\n\n"
                        )
            msg = (
                f"Saved {texts} text(s)"
                + (f" and {len(images)} image title(s)" if images else "")
                + f" to:\n{filename}"
            )
            self._show_popup(msg)
            try:
                os.startfile(str(BATCH_SAVE_DIR))
            except Exception as open_e:
                print(f"Warn: Open folder fail: {open_e}")
        except Exception as e:
            print(f"Error save batch: {e}")
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
                from tkinter import filedialog

                dialog.attributes("-topmost", False)
                filename = filedialog.asksaveasfilename(
                    initialdir=str(APP_DATA_DIR),
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
                    os.startfile(str(APP_DATA_DIR))
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
                    initialdir=str(APP_DATA_DIR),
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
                    self._filter_and_show()
                    self._show_popup(f"Restored clip group from:\n{filename}")
                    os.startfile(str(APP_DATA_DIR))
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
        dialog.title("Close Clippy")
        dialog.geometry("350x180")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.attributes("-topmost", True)
        dialog.configure(fg_color=MAIN_BG_COLOR)
        ctk.CTkLabel(
            dialog, text=f"History file location:\n{log_dir}\n\nClose Clippy now?"
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
        print("Clippy shutdown complete.")

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
                    "Show Clippy",
                    schedule(self._do_show_window),
                    default=True,
                    enabled=is_hid,
                ),
                item("Hide Clippy", schedule(self._hide_window), enabled=is_vis),
                item("Pause Capture", schedule(self._toggle_capture), enabled=is_r),
                item("Resume Capture", schedule(self._toggle_capture), enabled=is_p),
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
                item("Quit Clippy", schedule(self._prompt_close), enabled=self.running),
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

    def _show_popup(self, msg):
        """Schedules showing a message box in the main thread."""
        if self.running:
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
            print(f"Popup error: {e}")
            print(f"{APP_NAME} Msg: {msg}")


# --- Main Execution ---
if __name__ == "__main__":
    print(f"--- Starting {APP_NAME} --- {time.strftime('%Y-%m-d %H:%M:%S')}")
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
        print("DPI Awareness set (1)")
    except Exception as dpi_e:
        print(f"Info: DPI failed - {dpi_e}")
    try:
        app = Clippy()
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
