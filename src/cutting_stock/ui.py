import os
import sys
import csv
import re
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import List
from datetime import date

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib import colors
from reportlab.lib.utils import simpleSplit

from .models import CuttingJob
from .utils import (
    plan_cuts_for_job,
    calculate_efficiency,
    calculate_lost_material,
    PipeAssignment,
    build_results_summary,
    group_identical_pipe_assignments,
)


class CuttingStockUI:
    """
    Main UI controller for the Cutting Stock Optimizer.
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Cutting Stock Optimizer")
        self.root.geometry("1400x850")

        # Track which widget should receive mouse wheel scrolling
        self.active_mousewheel_widget = None

        # Track currently loaded/saved CSV file
        self.current_file_path = None

        self.is_dirty = False  # Tracks whether there are unsaved changes

        # Hidden save-file metadata dates
        self.creation_date = ""
        self.last_edit_date = ""

        # Store the latest computed result so it can be exported to PDF
        self.last_assignments = []
        self.last_new_pipe_count = 0
        self.last_efficiency = None
        self.last_kerf = 0
        self.last_summary_text = ""

        # Project information fields
        self.title_var = tk.StringVar()
        self.customer_var = tk.StringVar()
        self.notes_var = tk.StringVar()

        # Main input fields
        self.stock_len_var = tk.StringVar()
        self.kerf_var = tk.StringVar()
        self.include_leftovers_var = tk.BooleanVar(value=False)

        # Extra settings fields
        self.use_min_remainder_var = tk.BooleanVar(value=False)
        self.min_remainder_var = tk.StringVar()

        # Lost material settings
        self.calculate_lost_material_var = tk.BooleanVar(value=False)
        self.include_min_usable_length_var = tk.BooleanVar(value=False)
        self.min_usable_length_var = tk.StringVar()
        self.include_kerf_loss_var = tk.BooleanVar(value=False)

        # Tracks whether the leftover panel is currently visible
        self.leftovers_visible = True

        # Build UI
        self.setup_menu_bar()
        self.setup_main_layout()
        self.toggle_leftovers()
        self.toggle_min_remainder()
        self.toggle_lost_material_options()
        self.toggle_min_usable_length()

        # Global mouse wheel binding
        self.root.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        self.root.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.root.bind_all("<Button-5>", self._on_mousewheel_linux)

        # Keyboard shortcuts
        self.root.bind_all("<Control-s>", lambda event: self.save_plan())
        self.root.bind_all("<Control-S>", lambda event: self.new_save_plan())
        self.root.bind_all("<Control-n>", lambda event: self.new_plan())
        self.root.bind_all("<Control-o>", lambda event: self.load_plan())

    def setup_menu_bar(self):
        """
        Create the top dropdown menu bar.
        """
        menu_bar = tk.Menu(self.root)

        file_menu = tk.Menu(menu_bar, tearoff=0)
        file_menu.add_command(label="New Plan", command=self.new_plan, accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="Save Plan", command=self.save_plan, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As", command=self.new_save_plan, accelerator="Ctrl+Shift+S")
        file_menu.add_command(label="Open Plan", command=self.load_plan, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Export as PDF", command=self.export_results_pdf)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)

        edit_menu = tk.Menu(menu_bar, tearoff=0)
        edit_menu.add_command(label="Compute Cutting Plan", command=self.compute_plan)

        view_menu = tk.Menu(menu_bar, tearoff=0)
        view_menu.add_checkbutton(
            label="Use Leftover Pipes",
            variable=self.include_leftovers_var,
            command=self.toggle_leftovers
        )

        help_menu = tk.Menu(menu_bar, tearoff=0)
        help_menu.add_command(
            label="About",
            command=lambda: messagebox.showinfo(
                "About",
                "Cutting Stock Optimizer\nA pipe cutting optimization tool."
            )
        )

        menu_bar.add_cascade(label="File", menu=file_menu)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)
        menu_bar.add_cascade(label="View", menu=view_menu)
        menu_bar.add_cascade(label="Help", menu=help_menu)

        self.root.config(menu=menu_bar)

    def setup_main_layout(self):
        """
        Build the single-page resizable workspace.
        """

        # Main vertical splitter:
        # Top area = settings/cuts/leftovers/results
        # Bottom area = pipe visualization
        self.main_paned = tk.PanedWindow(
            self.root,
            orient=tk.VERTICAL,
            sashwidth=6,
            sashrelief="raised"
        )
        self.main_paned.pack(fill="both", expand=True)

        # Top horizontal splitter
        self.top_paned = tk.PanedWindow(
            self.main_paned,
            orient=tk.HORIZONTAL,
            sashwidth=6,
            sashrelief="raised"
        )
        self.main_paned.add(self.top_paned, minsize=300, stretch="always")

        # Bottom visualization panel
        self.visualization_panel = ttk.LabelFrame(
            self.main_paned,
            text="Pipe Visualization",
            padding=8
        )
        self.main_paned.add(self.visualization_panel, minsize=200, stretch="always")

        # Create panels
        self.setup_settings_panel()
        self.setup_cuts_panel()
        self.setup_leftovers_panel()
        self.setup_results_panel()
        self.setup_visualization_panel()

        # Fixed/natural-width panels on the left
        self.top_paned.add(self.settings_panel, minsize=260, width=300, stretch="never")
        self.top_paned.add(self.cuts_panel, minsize=360, width=390, stretch="never")
        self.top_paned.add(self.leftovers_panel, minsize=360, width=390, stretch="never")

        # Results panel gets the remaining space and expands with the window
        self.top_paned.add(self.results_panel, minsize=400, stretch="always")


    def setup_settings_panel(self):
        """
        Left-side settings panel with project data, pipe settings, and main action.
        The inside content is scrollable so extra settings still fit on smaller screens.
        """
        self.settings_panel = ttk.Frame(self.top_paned)

        self.settings_panel.rowconfigure(0, weight=1)
        self.settings_panel.columnconfigure(0, weight=1)

        # Scrollable canvas for settings content
        self.settings_canvas = tk.Canvas(self.settings_panel, highlightthickness=0)
        self.settings_canvas.grid(row=0, column=0, sticky="nsew")

        self.settings_scrollbar = ttk.Scrollbar(
            self.settings_panel,
            orient="vertical",
            command=self.settings_canvas.yview
        )
        self.settings_scrollbar.grid(row=0, column=1, sticky="ns")

        self.settings_canvas.configure(yscrollcommand=self.settings_scrollbar.set)

        self.settings_content = ttk.Frame(self.settings_canvas, padding=10)
        self.settings_window = self.settings_canvas.create_window(
            (0, 0),
            window=self.settings_content,
            anchor="nw"
        )

        self.settings_content.bind("<Configure>", self.on_settings_frame_configure)
        self.settings_canvas.bind("<Configure>", self.on_settings_canvas_configure)

        self._bind_mousewheel_target(self.settings_canvas, self.settings_canvas)
        self._bind_mousewheel_target(self.settings_content, self.settings_canvas)

        self.settings_content.columnconfigure(0, weight=0)
        self.settings_content.columnconfigure(1, weight=1)

        ttk.Label(
            self.settings_content,
            text="Input Settings",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, columnspan=2, pady=(10, 25))

        ttk.Label(self.settings_content, text="Title:").grid(
            row=1, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(self.settings_content, textvariable=self.title_var).grid(
            row=1, column=1, sticky="ew", pady=4
        )

        ttk.Label(self.settings_content, text="Customer:").grid(
            row=2, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(self.settings_content, textvariable=self.customer_var).grid(
            row=2, column=1, sticky="ew", pady=4
        )

        ttk.Label(self.settings_content, text="Additional Notes:").grid(
            row=3, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(self.settings_content, textvariable=self.notes_var).grid(
            row=3, column=1, sticky="ew", pady=4
        )

        ttk.Label(self.settings_content, text="Stock Pipe Length:").grid(
            row=4, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(self.settings_content, textvariable=self.stock_len_var).grid(
            row=4, column=1, sticky="ew", pady=4
        )

        ttk.Label(self.settings_content, text="Kerf:").grid(
            row=5, column=0, sticky="e", padx=(0, 8), pady=4
        )
        ttk.Entry(self.settings_content, textvariable=self.kerf_var).grid(
            row=5, column=1, sticky="ew", pady=4
        )

        # Use Leftover Pipes moved above minimum remainder
        ttk.Checkbutton(
            self.settings_content,
            text="Use Leftover Pipes",
            variable=self.include_leftovers_var,
            command=self.toggle_leftovers
        ).grid(row=6, column=0, columnspan=2, sticky="w", pady=(12, 4))

        ttk.Checkbutton(
            self.settings_content,
            text="Use a minimum remainder",
            variable=self.use_min_remainder_var,
            command=self.toggle_min_remainder
        ).grid(row=7, column=0, columnspan=2, sticky="w", pady=(12, 4))

        self.min_remainder_frame = ttk.Frame(self.settings_content)

        ttk.Label(
            self.min_remainder_frame,
            text="Minimum Remainder:"
        ).pack(side="left", padx=(0, 8))

        ttk.Entry(
            self.min_remainder_frame,
            textvariable=self.min_remainder_var,
            width=10
        ).pack(side="left")

        self.min_remainder_frame.grid(row=8, column=0, columnspan=2, sticky="w", padx=(25, 0), pady=4)

        # Calculate lost material section
        ttk.Checkbutton(
            self.settings_content,
            text="Calculate lost material",
            variable=self.calculate_lost_material_var,
            command=self.toggle_lost_material_options
        ).grid(row=9, column=0, columnspan=2, sticky="w", pady=(18, 4))

        self.lost_material_options_frame = ttk.Frame(self.settings_content)

        ttk.Checkbutton(
            self.lost_material_options_frame,
            text="Include minimum usable length",
            variable=self.include_min_usable_length_var,
            command=self.toggle_min_usable_length
        ).grid(row=0, column=0, columnspan=2, sticky="w", padx=(25, 0), pady=4)

        self.min_usable_length_frame = ttk.Frame(self.lost_material_options_frame)

        ttk.Label(
            self.min_usable_length_frame,
            text="Minimum usable length:"
        ).pack(side="left", padx=(0, 8))

        ttk.Entry(
            self.min_usable_length_frame,
            textvariable=self.min_usable_length_var,
            width=10
        ).pack(side="left")

        self.min_usable_length_frame.grid(row=1, column=0, columnspan=2, sticky="w", padx=(45, 0), pady=4)

        ttk.Checkbutton(
            self.lost_material_options_frame,
            text="Include kerf loss",
            variable=self.include_kerf_loss_var
        ).grid(row=2, column=0, columnspan=2, sticky="w", padx=(25, 0), pady=4)

        self.lost_material_options_frame.grid(row=10, column=0, columnspan=2, sticky="w", pady=4)

        ttk.Button(
            self.settings_content,
            text="Compute Cutting Plan",
            command=self.compute_plan
        ).grid(row=11, column=0, columnspan=2, pady=25)



###### SEGMENT 2
    def setup_cuts_panel(self):
        """
        Required cuts panel.
        """
        self.cuts_panel = ttk.Frame(self.top_paned, padding=8)
        self.cuts_panel.columnconfigure(0, weight=1)
        self.cuts_panel.rowconfigure(2, weight=1)

        ttk.Label(
            self.cuts_panel,
            text="Required Cuts",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, pady=(0, 10))

        self.cuts_header_frame = ttk.Frame(self.cuts_panel)
        self.cuts_header_frame.grid(row=1, column=0, sticky="w", pady=(0, 5))

        ttk.Label(self.cuts_header_frame, text="Label", width=12, anchor="center").grid(row=0, column=0, padx=2)
        ttk.Label(self.cuts_header_frame, text="Length", width=8, anchor="center").grid(row=0, column=1, padx=2)
        ttk.Label(self.cuts_header_frame, text="Quantity", width=8, anchor="center").grid(row=0, column=2, padx=2)
        ttk.Label(self.cuts_header_frame, text="", width=8).grid(row=0, column=3, padx=2)

        self.cuts_scroll_container = ttk.Frame(self.cuts_panel, borderwidth=1, relief="solid")
        self.cuts_scroll_container.grid(row=2, column=0, sticky="nsew")

        self.cuts_scroll_container.rowconfigure(0, weight=1)
        self.cuts_scroll_container.columnconfigure(0, weight=1)

        self.cuts_canvas = tk.Canvas(self.cuts_scroll_container, highlightthickness=0)
        self.cuts_canvas.grid(row=0, column=0, sticky="nsew")

        self.cuts_scrollbar = ttk.Scrollbar(
            self.cuts_scroll_container,
            orient="vertical",
            command=self.cuts_canvas.yview
        )
        self.cuts_scrollbar.grid(row=0, column=1, sticky="ns")

        self.cuts_canvas.configure(yscrollcommand=self.cuts_scrollbar.set)

        self.cuts_frame = ttk.Frame(self.cuts_canvas)
        self.cuts_window = self.cuts_canvas.create_window((0, 0), window=self.cuts_frame, anchor="nw")

        self.cuts_frame.bind("<Configure>", self.on_cuts_frame_configure)
        self.cuts_canvas.bind("<Configure>", self.on_cuts_canvas_configure)

        self._bind_mousewheel_target(self.cuts_canvas, self.cuts_canvas)
        self._bind_mousewheel_target(self.cuts_frame, self.cuts_canvas)

        self.cuts_rows: List[List[tk.Entry]] = []
        self.add_cuts_row()
        self.update_cuts_add_button()

    def setup_leftovers_panel(self):
        """
        Leftover pipes panel.
        """
        self.leftovers_panel = ttk.Frame(self.top_paned, padding=8)
        self.leftovers_panel.columnconfigure(0, weight=1)
        self.leftovers_panel.rowconfigure(2, weight=1)

        ttk.Label(
            self.leftovers_panel,
            text="Leftover Pipes",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, pady=(0, 10))

        self.leftovers_header_frame = ttk.Frame(self.leftovers_panel)
        self.leftovers_header_frame.grid(row=1, column=0, sticky="w", pady=(0, 5))

        ttk.Label(self.leftovers_header_frame, text="Label", width=12, anchor="center").grid(row=0, column=0, padx=2)
        ttk.Label(self.leftovers_header_frame, text="Length", width=8, anchor="center").grid(row=0, column=1, padx=2)
        ttk.Label(self.leftovers_header_frame, text="Quantity", width=8, anchor="center").grid(row=0, column=2, padx=2)
        ttk.Label(self.leftovers_header_frame, text="", width=8).grid(row=0, column=3, padx=2)

        self.leftovers_scroll_container = ttk.Frame(self.leftovers_panel, borderwidth=1, relief="solid")
        self.leftovers_scroll_container.grid(row=2, column=0, sticky="nsew")

        self.leftovers_scroll_container.rowconfigure(0, weight=1)
        self.leftovers_scroll_container.columnconfigure(0, weight=1)

        self.leftovers_canvas = tk.Canvas(self.leftovers_scroll_container, highlightthickness=0)
        self.leftovers_canvas.grid(row=0, column=0, sticky="nsew")

        self.leftovers_scrollbar = ttk.Scrollbar(
            self.leftovers_scroll_container,
            orient="vertical",
            command=self.leftovers_canvas.yview
        )
        self.leftovers_scrollbar.grid(row=0, column=1, sticky="ns")

        self.leftovers_canvas.configure(yscrollcommand=self.leftovers_scrollbar.set)

        self.leftovers_frame = ttk.Frame(self.leftovers_canvas)
        self.leftovers_window = self.leftovers_canvas.create_window((0, 0), window=self.leftovers_frame, anchor="nw")

        self.leftovers_frame.bind("<Configure>", self.on_leftovers_frame_configure)
        self.leftovers_canvas.bind("<Configure>", self.on_leftovers_canvas_configure)

        self._bind_mousewheel_target(self.leftovers_canvas, self.leftovers_canvas)
        self._bind_mousewheel_target(self.leftovers_frame, self.leftovers_canvas)

        self.leftovers_rows: List[List[tk.Entry]] = []
        self.add_leftovers_row()
        self.update_leftovers_add_button()

    def setup_results_panel(self):
        """
        Results text panel.
        """
        self.results_panel = ttk.LabelFrame(self.top_paned, text="Results", padding=8)
        self.results_panel.rowconfigure(0, weight=1)
        self.results_panel.columnconfigure(0, weight=1)

        self.results_text = tk.Text(
            self.results_panel,
            wrap="word",
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 10),
            height=1
        )
        self.results_text.grid(row=0, column=0, sticky="nsew")

        results_scroll = ttk.Scrollbar(self.results_panel, orient="vertical", command=self.results_text.yview)
        results_scroll.grid(row=0, column=1, sticky="ns")
        self.results_text.configure(yscrollcommand=results_scroll.set)

        self._bind_mousewheel_target(self.results_text, self.results_text)

    def setup_visualization_panel(self):
        """
        Bottom pipe visualization panel.
        """
        self.visualization_panel.rowconfigure(0, weight=1)
        self.visualization_panel.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            self.visualization_panel,
            bg="white",
            highlightthickness=0,
            height=1
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        canvas_vscroll = ttk.Scrollbar(self.visualization_panel, orient="vertical", command=self.canvas.yview)
        canvas_vscroll.grid(row=0, column=1, sticky="ns")

        canvas_hscroll = ttk.Scrollbar(self.visualization_panel, orient="horizontal", command=self.canvas.xview)
        canvas_hscroll.grid(row=1, column=0, sticky="ew")

        self.canvas.configure(
            yscrollcommand=canvas_vscroll.set,
            xscrollcommand=canvas_hscroll.set
        )

        self._bind_mousewheel_target(self.canvas, self.canvas)

    # =========================
    # MOUSE WHEEL SUPPORT
    # =========================

    def _bind_mousewheel_target(self, widget, scroll_target):
        widget.bind("<Enter>", lambda e, target=scroll_target: self._set_active_mousewheel(target))

    def _set_active_mousewheel(self, target):
        self.active_mousewheel_widget = target

    def _on_mousewheel_windows(self, event):
        if self.active_mousewheel_widget is None:
            return

        try:
            delta = int(-1 * (event.delta / 120))
            if delta != 0:
                self.active_mousewheel_widget.yview_scroll(delta, "units")
        except Exception:
            pass

    def _on_mousewheel_linux(self, event):
        if self.active_mousewheel_widget is None:
            return

        try:
            if event.num == 4:
                self.active_mousewheel_widget.yview_scroll(-1, "units")
            elif event.num == 5:
                self.active_mousewheel_widget.yview_scroll(1, "units")
        except Exception:
            pass

    # =========================
    # SCROLLABLE PANEL HELPERS
    # =========================

    def on_cuts_frame_configure(self, event=None):
        self.cuts_canvas.configure(scrollregion=self.cuts_canvas.bbox("all"))


    def on_cuts_canvas_configure(self, event):
        self.cuts_frame.update_idletasks()
        required_width = self.cuts_frame.winfo_reqwidth()
        self.cuts_canvas.itemconfig(self.cuts_window, width=max(event.width, required_width))


    def on_leftovers_frame_configure(self, event=None):
        self.leftovers_canvas.configure(scrollregion=self.leftovers_canvas.bbox("all"))


    def on_leftovers_canvas_configure(self, event):
        self.leftovers_frame.update_idletasks()
        required_width = self.leftovers_frame.winfo_reqwidth()
        self.leftovers_canvas.itemconfig(self.leftovers_window, width=max(event.width, required_width))


    def on_settings_frame_configure(self, event=None):
        self.settings_canvas.configure(scrollregion=self.settings_canvas.bbox("all"))


    def on_settings_canvas_configure(self, event):
        self.settings_canvas.itemconfig(self.settings_window, width=event.width)


    def toggle_lost_material_options(self):
        """
        Show lost material options only when Calculate lost material is checked.
        """
        if self.calculate_lost_material_var.get():
            self.lost_material_options_frame.grid()
        else:
            self.lost_material_options_frame.grid_remove()


    def toggle_min_usable_length(self):
        """
        Show Minimum usable length input only when enabled.
        """
        if self.include_min_usable_length_var.get():
            self.min_usable_length_frame.grid()
        else:
            self.min_usable_length_frame.grid_remove()


    def toggle_leftovers(self):
        """
        Show or hide the leftover pipes panel in the resizable workspace.
        """
        if self.include_leftovers_var.get():
            if not self.leftovers_visible:
                self.top_paned.forget(self.results_panel)
                self.top_paned.add(self.leftovers_panel, minsize=360, width=390, stretch="never")
                self.top_paned.add(self.results_panel, minsize=400, stretch="always")
                self.leftovers_visible = True
        else:
            if self.leftovers_visible:
                self.top_paned.forget(self.leftovers_panel)
                self.leftovers_visible = False

    def toggle_min_remainder(self):
        """
        Show Minimum Remainder only when Use a minimum remainder is checked.
        """
        if self.use_min_remainder_var.get():
            self.min_remainder_frame.grid()
        else:
            self.min_remainder_frame.grid_remove()

    # =========================

    def add_cuts_row(self):
        row_num = len(self.cuts_rows)

        label_entry = ttk.Entry(self.cuts_frame, width=12)
        length_entry = ttk.Entry(self.cuts_frame, width=8)
        qty_entry = ttk.Entry(self.cuts_frame, width=8)

        delete_btn = tk.Button(
            self.cuts_frame,
            text="Delete",
            fg="white",
            bg="red",
            width=8,
            font=("TkDefaultFont", 9, "bold"),
            command=lambda: self.remove_cuts_row(row_num)
        )

        label_entry.grid(row=row_num, column=0, padx=2, pady=2, sticky="w")
        length_entry.grid(row=row_num, column=1, padx=2, pady=2, sticky="w")
        qty_entry.grid(row=row_num, column=2, padx=2, pady=2, sticky="w")
        delete_btn.grid(row=row_num, column=3, padx=2, pady=2, sticky="w")

        qty_entry.bind("<Tab>", lambda e: self.tab_to_next_row(self.cuts_rows, row_num, self.add_cuts_row))

        self.cuts_rows.append([label_entry, length_entry, qty_entry])
        self.update_cuts_add_button()

    def add_leftovers_row(self):
        row_num = len(self.leftovers_rows)

        label_entry = ttk.Entry(self.leftovers_frame, width=12)
        length_entry = ttk.Entry(self.leftovers_frame, width=8)
        qty_entry = ttk.Entry(self.leftovers_frame, width=8)

        delete_btn = tk.Button(
            self.leftovers_frame,
            text="Delete",
            fg="white",
            bg="red",
            width=8,
            font=("TkDefaultFont", 9, "bold"),
            command=lambda: self.remove_leftovers_row(row_num)
        )

        label_entry.grid(row=row_num, column=0, padx=2, pady=2, sticky="w")
        length_entry.grid(row=row_num, column=1, padx=2, pady=2, sticky="w")
        qty_entry.grid(row=row_num, column=2, padx=2, pady=2, sticky="w")
        delete_btn.grid(row=row_num, column=3, padx=2, pady=2, sticky="w")

        qty_entry.bind("<Tab>", lambda e: self.tab_to_next_row(self.leftovers_rows, row_num, self.add_leftovers_row))

        self.leftovers_rows.append([label_entry, length_entry, qty_entry])
        self.update_leftovers_add_button()

    def tab_to_next_row(self, rows_list, current_row, add_func):
        if current_row + 1 >= len(rows_list):
            add_func()

        next_row = current_row + 1
        if next_row < len(rows_list):
            rows_list[next_row][0].focus_set()

        return "break"

    def remove_cuts_row(self, row_idx):
        if row_idx < len(self.cuts_rows):
            for widget in self.cuts_rows[row_idx]:
                widget.destroy()

            del self.cuts_rows[row_idx]

            for widget in list(self.cuts_frame.winfo_children()):
                if isinstance(widget, tk.Button) and widget != getattr(self, 'cuts_add_btn', None):
                    widget.destroy()

            for i, widgets in enumerate(self.cuts_rows):
                for j, widget in enumerate(widgets):
                    widget.grid(row=i, column=j, padx=2, pady=2, sticky="w")

                delete_btn = tk.Button(
                    self.cuts_frame,
                    text="Delete",
                    fg="white",
                    bg="red",
                    width=8,
                    font=("TkDefaultFont", 9, "bold"),
                    command=lambda idx=i: self.remove_cuts_row(idx)
                )
                delete_btn.grid(row=i, column=3, padx=2, pady=2, sticky="w")

            self.update_cuts_add_button()

    def remove_leftovers_row(self, row_idx):
        if row_idx < len(self.leftovers_rows):
            for widget in self.leftovers_rows[row_idx]:
                widget.destroy()

            del self.leftovers_rows[row_idx]

            for widget in list(self.leftovers_frame.winfo_children()):
                if isinstance(widget, tk.Button) and widget != getattr(self, 'leftovers_add_btn', None):
                    widget.destroy()

            for i, widgets in enumerate(self.leftovers_rows):
                for j, widget in enumerate(widgets):
                    widget.grid(row=i, column=j, padx=2, pady=2, sticky="w")

                delete_btn = tk.Button(
                    self.leftovers_frame,
                    text="Delete",
                    fg="white",
                    bg="red",
                    width=8,
                    font=("TkDefaultFont", 9, "bold"),
                    command=lambda idx=i: self.remove_leftovers_row(idx)
                )
                delete_btn.grid(row=i, column=3, padx=2, pady=2, sticky="w")

            self.update_leftovers_add_button()

    def update_cuts_add_button(self):
        if hasattr(self, 'cuts_add_btn') and self.cuts_add_btn.winfo_exists():
            self.cuts_add_btn.destroy()

        row_num = len(self.cuts_rows)

        self.cuts_add_btn = tk.Button(
            self.cuts_frame,
            text="+ Add Row",
            bg="lightgreen",
            font=("Arial", 10, "bold"),
            command=self.add_cuts_row
        )
        self.cuts_add_btn.grid(row=row_num, column=0, columnspan=4, pady=8)

    def update_leftovers_add_button(self):
        if hasattr(self, 'leftovers_add_btn') and self.leftovers_add_btn.winfo_exists():
            self.leftovers_add_btn.destroy()

        row_num = len(self.leftovers_rows)

        self.leftovers_add_btn = tk.Button(
            self.leftovers_frame,
            text="+ Add Row",
            bg="lightgreen",
            font=("Arial", 10, "bold"),
            command=self.add_leftovers_row
        )
        self.leftovers_add_btn.grid(row=row_num, column=0, columnspan=4, pady=8)

    def get_today_string(self):
        return date.today().strftime("%Y-%m-%d")

    def make_default_filename(self):
        today = self.get_today_string()
        title = self.title_var.get().strip()

        if title:
            safe_title = re.sub(r'[<>:"/\\|?*]', "", title)
            safe_title = safe_title.replace(" ", "_")
            return f"{today}_{safe_title}.csv"

        return f"{today}_cutting_plan.csv"


    def new_plan(self):
        """
        Clear the current workspace and return the app to a fresh startup state.
        If there is an open/edited plan, ask whether to save first.
        """
        if self.is_dirty or self.current_file_path:
            answer = messagebox.askyesnocancel(
                "New Plan",
                "Do you want to save the current plan before creating a new one?"
            )

            if answer is None:
                return

            if answer is True:
                self.save_plan()

        self.current_file_path = None
        self.creation_date = ""
        self.last_edit_date = ""

        self.title_var.set("")
        self.customer_var.set("")
        self.notes_var.set("")
        self.stock_len_var.set("")
        self.kerf_var.set("")

        self.include_leftovers_var.set(False)
        self.use_min_remainder_var.set(False)
        self.min_remainder_var.set("")

        self.calculate_lost_material_var.set(False)
        self.include_min_usable_length_var.set(False)
        self.min_usable_length_var.set("")
        self.include_kerf_loss_var.set(False)

        while len(self.cuts_rows) > 0:
            self.remove_cuts_row(0)

        while len(self.leftovers_rows) > 0:
            self.remove_leftovers_row(0)

        self.add_cuts_row()
        self.add_leftovers_row()

        self.toggle_leftovers()
        self.toggle_min_remainder()
        self.toggle_lost_material_options()
        self.toggle_min_usable_length()

        self.last_assignments = []
        self.last_new_pipe_count = 0
        self.last_efficiency = None
        self.last_kerf = 0
        self.last_summary_text = ""

        self.results_text.delete(1.0, tk.END)
        self.canvas.delete("all")
        self.canvas.configure(scrollregion=(0, 0, 0, 0))

        self.is_dirty = False


    def new_save_plan(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=self.make_default_filename()
        )

        if not file_path:
            return

        try:
            self.write_plan_to_file(file_path, new_file=True)
            messagebox.showinfo("Success", f"Plan saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save plan: {e}")


## =========================


    def save_plan(self):
        if not self.current_file_path:
            self.new_save_plan()
            return

        try:
            self.write_plan_to_file(self.current_file_path, new_file=False)
            messagebox.showinfo("Success", f"Plan saved to {self.current_file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save plan: {e}")

    def write_plan_to_file(self, file_path, new_file=False):
        today = self.get_today_string()

        if new_file or not self.creation_date:
            self.creation_date = today

        self.last_edit_date = today

        leftovers = self.parse_grid_input(self.leftovers_rows, "Left Over") if self.include_leftovers_var.get() else []
        cuts = self.parse_grid_input(self.cuts_rows, "SEG")

        with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
            writer = csv.writer(csvfile)

            writer.writerow(["INFORMATION:"])
            writer.writerow(["Title", self.title_var.get()])
            writer.writerow(["Creation Date", self.creation_date])
            writer.writerow(["Last Edit Date", self.last_edit_date])
            writer.writerow(["Customer", self.customer_var.get()])
            writer.writerow(["Additional Notes", self.notes_var.get()])
            writer.writerow(["Stock Pipe Length", self.stock_len_var.get()])
            writer.writerow(["Kerf", self.kerf_var.get()])
            writer.writerow([])

            writer.writerow(["SETTINGS:"])
            writer.writerow(["Use Leftover Pipes", self.include_leftovers_var.get()])
            writer.writerow(["Use a minimum remainder", self.use_min_remainder_var.get()])
            writer.writerow(["Minimum Remainder", self.min_remainder_var.get()])
            writer.writerow(["Calculate lost material", self.calculate_lost_material_var.get()])
            writer.writerow(["Include minimum usable length", self.include_min_usable_length_var.get()])
            writer.writerow(["Minimum usable length", self.min_usable_length_var.get()])
            writer.writerow(["Include kerf loss", self.include_kerf_loss_var.get()])
            writer.writerow([])

            writer.writerow(["LEFTOVER PIPES:"])
            writer.writerow(["Quantity", "Length", "Label"])
            for label, length, qty in leftovers:
                writer.writerow([qty, length, label])
            writer.writerow([])

            writer.writerow(["NEW CUTS:"])
            writer.writerow(["Quantity", "Length", "Label"])
            for label, length, qty in cuts:
                writer.writerow([qty, length, label])

        self.current_file_path = file_path
        self.is_dirty = False


    def load_plan(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as csvfile:
                reader = list(csv.reader(csvfile))

            cuts = []
            leftovers = []

            self.creation_date = ""
            self.last_edit_date = ""

            section = None

            for row in reader:
                if not row:
                    continue

                first = row[0].strip()

                if first == "INFORMATION:":
                    section = "info"
                    continue
                elif first == "SETTINGS:":
                    section = "settings"
                    continue
                elif first == "LEFTOVER PIPES:":
                    section = "leftovers"
                    continue
                elif first == "NEW CUTS:":
                    section = "cuts"
                    continue

                if first in ["Quantity"]:
                    continue

                if section == "info" and len(row) >= 2:
                    key, value = row[0], row[1]

                    if key == "Title":
                        self.title_var.set(value)
                    elif key == "Creation Date":
                        self.creation_date = value
                    elif key == "Last Edit Date":
                        self.last_edit_date = value
                    elif key == "Customer":
                        self.customer_var.set(value)
                    elif key == "Additional Notes":
                        self.notes_var.set(value)
                    elif key == "Stock Pipe Length":
                        self.stock_len_var.set(value)
                    elif key == "Kerf":
                        self.kerf_var.set(value)

                elif section == "settings" and len(row) >= 2:
                    key, value = row[0], row[1]

                    if key == "Use Leftover Pipes":
                        self.include_leftovers_var.set(value.lower() == "true")
                    elif key == "Use a minimum remainder":
                        self.use_min_remainder_var.set(value.lower() == "true")
                    elif key == "Minimum Remainder":
                        self.min_remainder_var.set(value)
                    elif key == "Calculate lost material":
                        self.calculate_lost_material_var.set(value.lower() == "true")
                    elif key == "Include minimum usable length":
                        self.include_min_usable_length_var.set(value.lower() == "true")
                    elif key == "Minimum usable length":
                        self.min_usable_length_var.set(value)
                    elif key == "Include kerf loss":
                        self.include_kerf_loss_var.set(value.lower() == "true")

#               #
                elif section == "leftovers":
                    if len(row) < 2:
                        continue

                    qty = row[0].strip()
                    length = row[1].strip()
                    label = row[2].strip() if len(row) >= 3 else ""

                    if not qty or not length:
                        continue

                    leftovers.append((label, int(length), int(qty)))

                elif section == "cuts":
                    if len(row) < 2:
                        continue

                    qty = row[0].strip()
                    length = row[1].strip()
                    label = row[2].strip() if len(row) >= 3 else ""

                    if not qty or not length:
                        continue

                    cuts.append((label, int(length), int(qty)))
#               #

            while len(self.cuts_rows) > 0:
                self.remove_cuts_row(0)
            while len(self.leftovers_rows) > 0:
                self.remove_leftovers_row(0)

            for label, length, qty in cuts:
                self.add_cuts_row()
                row = self.cuts_rows[-1]
                row[0].insert(0, label)
                row[1].insert(0, str(length))
                row[2].insert(0, str(qty))

            for label, length, qty in leftovers:
                self.add_leftovers_row()
                row = self.leftovers_rows[-1]
                row[0].insert(0, label)
                row[1].insert(0, str(length))
                row[2].insert(0, str(qty))

            self.toggle_leftovers()
            self.toggle_min_remainder()
            self.toggle_lost_material_options()
            self.toggle_min_usable_length()

            self.current_file_path = file_path

            self.last_assignments = []
            self.last_new_pipe_count = 0
            self.last_efficiency = None
            self.last_summary_text = ""

            self.results_text.delete(1.0, tk.END)
            self.canvas.delete("all")
            self.canvas.configure(scrollregion=(0, 0, 0, 0))

            self.is_dirty = False

            messagebox.showinfo("Success", f"Plan loaded from {file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to load plan: {e}")


# =========================


    def parse_grid_input(self, rows_list, default_prefix):
        result = []
        label_counter = 1

        for row in rows_list:
            label = row[0].get().strip()
            length_str = row[1].get().strip()
            qty_str = row[2].get().strip()

            if not length_str:
                continue

            try:
                length = int(length_str)
                qty = int(qty_str) if qty_str else 1

                if label == "":
                    label = f"{default_prefix}_{label_counter:02d}"
                    label_counter += 1

                result.append((label, length, qty))
            except ValueError:
                raise ValueError(f"Invalid length or quantity in row: {label}, {length_str}, {qty_str}")

        return result

    def compute_plan(self):
        try:
            cut_req = self.parse_grid_input(self.cuts_rows, "SEG")
            leftovers = []
            if self.include_leftovers_var.get():
                leftovers = self.parse_grid_input(self.leftovers_rows, "Left Over")

            stock_len_str = self.stock_len_var.get().strip()
            kerf_str = self.kerf_var.get().strip()

            if not stock_len_str or not kerf_str:
                messagebox.showerror("Error", "Stock Pipe Length and Kerf are required")
                return

            stock_len = int(stock_len_str)
            kerf = int(kerf_str)

            job = CuttingJob(
                cut_requirements=cut_req,
                stock_pipe_length=stock_len,
                leftover_pipes=leftovers,
                kerf=kerf,
                include_leftovers=self.include_leftovers_var.get(),
            )

            #
            use_minimum_remainder = self.use_min_remainder_var.get()
            minimum_remainder = 0

            if use_minimum_remainder:
                min_remainder_str = self.min_remainder_var.get().strip()

                if not min_remainder_str:
                    messagebox.showerror("Error", "Minimum Remainder is required when enabled")
                    return

                try:
                    minimum_remainder = int(min_remainder_str)
                except ValueError:
                    messagebox.showerror("Error", "Minimum Remainder must be a number")
                    return

                if minimum_remainder < 0:
                    messagebox.showerror("Error", "Minimum Remainder cannot be negative")
                    return

            assignments, new_pipe_count = plan_cuts_for_job(
                job,
                use_minimum_remainder=use_minimum_remainder,
                minimum_remainder=minimum_remainder
            )
            #
            efficiency = calculate_efficiency(assignments)
            #
            lost_material = None

            if self.calculate_lost_material_var.get():
                minimum_usable_length = 0

                if self.include_min_usable_length_var.get():
                    min_usable_str = self.min_usable_length_var.get().strip()

                    if not min_usable_str:
                        messagebox.showerror("Error", "Minimum usable length is required when enabled")
                        return

                    try:
                        minimum_usable_length = int(min_usable_str)
                    except ValueError:
                        messagebox.showerror("Error", "Minimum usable length must be a number")
                        return

                    if minimum_usable_length < 0:
                        messagebox.showerror("Error", "Minimum usable length cannot be negative")
                        return

                lost_material = calculate_lost_material(
                    assignments,
                    kerf,
                    include_minimum_usable_length=self.include_min_usable_length_var.get(),
                    minimum_usable_length=minimum_usable_length,
                    include_kerf_loss=self.include_kerf_loss_var.get()
                )
            #

            self.last_assignments = assignments
            self.last_new_pipe_count = new_pipe_count
            self.last_efficiency = efficiency
            self.last_kerf = kerf
            self.last_summary_text = build_results_summary(assignments, new_pipe_count, efficiency, kerf)
            #
            if lost_material is not None:
                self.last_summary_text += "\n\n--- Lost Material ---"
                self.last_summary_text += f"\nTotal lost material: {lost_material['total_lost_material']} mm"

                if self.include_min_usable_length_var.get():
                    self.last_summary_text += f"\nUnusable remainder: {lost_material['unusable_remainder']} mm"

                if self.include_kerf_loss_var.get():
                    self.last_summary_text += f"\nKerf loss: {lost_material['kerf_loss']} mm"
            #

            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, self.last_summary_text)

            self.visualize_pipes(assignments, kerf)

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def visualize_pipes(self, assignments: list[PipeAssignment], kerf: int):
        self.canvas.delete("all")
        grouped_pipes = group_identical_pipe_assignments(assignments)

        y_offset = 20
        max_width = 1600

        max_pipe_length = max((pipe.original_length for pipe in assignments), default=5000)
        scale = max_width / max(5000, max_pipe_length)

        colors_list = ["red", "blue", "green", "orange", "purple", "brown", "pink", "gray", "cyan", "magenta"]
        font_label = ("Segoe UI", 8, "bold")
        font_pipe = ("Segoe UI", 9, "bold")

        for pipe, count in grouped_pipes:

            pipe_right = 10 + int(pipe.original_length * scale)

            self.canvas.create_text(
                10, y_offset - 10,
                text=f"Pipe {pipe.id} x{count} ({pipe.source})",
                anchor="w",
                font=font_pipe
            )

            self.canvas.create_rectangle(
                10, y_offset,
                pipe_right, y_offset + 30,
                outline="black",
                fill="lightgray"
            )

            cut_x = 10.0

            for i, cut in enumerate(pipe.cuts):
                cut_width = cut.length * scale

                if i == len(pipe.cuts) - 1:
                    current_end = cut_x + cut_width
                    if pipe.remaining_length == 0 and current_end < pipe_right:
                        cut_width += (pipe_right - current_end)

                color = colors_list[i % len(colors_list)]

                self.canvas.create_rectangle(
                    int(cut_x), y_offset,
                    int(cut_x + cut_width), y_offset + 30,
                    fill=color,
                    outline="black"
                )

                text_label = f"{cut.id}({cut.length})"
                self.canvas.create_text(
                    int(cut_x + cut_width / 2),
                    y_offset + 15,
                    text=text_label,
                    fill="white",
                    font=font_label,
                    width=max(int(cut_width) - 4, 1)
                )

                if i < len(pipe.cuts) - 1:
                    cut_x += cut_width + (kerf * scale)
                else:
                    cut_x += cut_width

            y_offset += 50

        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def export_results_pdf(self):
        if not self.last_assignments or self.last_efficiency is None:
            messagebox.showerror("Error", "No cutting plan has been computed yet.")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile="cutting_plan.pdf"
        )

        if not file_path:
            return

        try:
            pdf = pdf_canvas.Canvas(file_path, pagesize=A4)
            page_width, page_height = A4

            margin = 40
            y = page_height - margin

            pdf.setFont("Helvetica-Bold", 14)
            pdf.drawString(margin, y, "Cutting Stock Optimizer - Cutting Plan")
            y -= 25

            if self.title_var.get().strip():
                pdf.setFont("Helvetica", 10)
                pdf.drawString(margin, y, f"Title: {self.title_var.get().strip()}")
                y -= 14

            if self.customer_var.get().strip():
                pdf.setFont("Helvetica", 10)
                pdf.drawString(margin, y, f"Customer: {self.customer_var.get().strip()}")
                y -= 14

            if self.notes_var.get().strip():
                pdf.setFont("Helvetica", 10)
                pdf.drawString(margin, y, f"Notes: {self.notes_var.get().strip()}")
                y -= 14

            y -= 10

            pdf.setFont("Helvetica", 10)
            line_height = 14
            max_text_width = page_width - (2 * margin)

            for line in self.last_summary_text.splitlines():
                wrapped_lines = simpleSplit(line, "Helvetica", 10, max_text_width)

                if not wrapped_lines:
                    wrapped_lines = [""]

                for wrapped_line in wrapped_lines:
                    if y < 220:
                        pdf.showPage()
                        y = page_height - margin
                        pdf.setFont("Helvetica", 10)

                    pdf.drawString(margin, y, wrapped_line)
                    y -= line_height

            y -= 20

            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(margin, y, "Pipe Visualization")
            y -= 20

            available_width = page_width - (2 * margin)

            pdf_colors = [
                colors.red,
                colors.blue,
                colors.green,
                colors.orange,
                colors.purple,
                colors.brown,
                colors.pink,
                colors.gray,
                colors.cyan,
                colors.magenta,
            ]

            max_pipe_length = max(
                (pipe.original_length for pipe in self.last_assignments if pipe.cuts),
                default=5000
            )
            scale = available_width / max(5000, max_pipe_length)

            bar_height = 18
            pipe_spacing = 38

            for pipe_index, pipe in enumerate(self.last_assignments, start=1):
                if not pipe.cuts:
                    continue

                if y < 80:
                    pdf.showPage()
                    y = page_height - margin
                    pdf.setFont("Helvetica-Bold", 12)
                    pdf.setFillColor(colors.black)
                    pdf.drawString(margin, y, "Pipe Visualization (continued)")
                    y -= 20

                pdf.setFont("Helvetica-Bold", 10)
                pdf.setFillColor(colors.black)
                pdf.drawString(margin, y, f"Pipe {pipe_index} ({pipe.source})")
                y -= 12

                pipe_x = margin
                pipe_y = y - bar_height
                pipe_width = pipe.original_length * scale

                pdf.setStrokeColor(colors.black)
                pdf.setFillColor(colors.lightgrey)
                pdf.rect(pipe_x, pipe_y, pipe_width, bar_height, stroke=1, fill=1)

                cut_x = pipe_x

                for i, cut in enumerate(pipe.cuts):
                    cut_width = cut.length * scale
                    fill_color = pdf_colors[i % len(pdf_colors)]

                    pdf.setFillColor(fill_color)
                    pdf.rect(cut_x, pipe_y, cut_width, bar_height, stroke=1, fill=1)

                    if cut_width > 45:
                        pdf.setFillColor(colors.white)
                        pdf.setFont("Helvetica-Bold", 8)
                        pdf.drawCentredString(
                            cut_x + (cut_width / 2),
                            pipe_y + 5,
                            f"{cut.id}({cut.length})"
                        )

                    if i < len(pipe.cuts) - 1:
                        cut_x += cut_width + (self.last_kerf * scale)
                    else:
                        cut_x += cut_width

                y -= pipe_spacing

            pdf.save()
            messagebox.showinfo("Success", f"PDF exported to:\n{file_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to export PDF: {e}")


def resource_path(relative_path):
    """
    Get absolute path to resource, works for dev and for PyInstaller.
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)

            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "Martin.CuttingStockOptimizer.1"
            )
        except Exception:
            pass

    root = tk.Tk()

    try:
        ico_path = resource_path(os.path.join("assets", "CSO_icon.ico"))
        png_path = resource_path(os.path.join("assets", "CSO_icon.png"))

        if sys.platform == "win32":
            root.iconbitmap(ico_path)

        icon_image = tk.PhotoImage(file=png_path)
        root.iconphoto(True, icon_image)
        root._icon_image = icon_image

    except Exception as e:
        print(f"Warning: Could not load icon: {e}")

    app = CuttingStockUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()