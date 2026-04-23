import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import csv
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas as pdf_canvas
from reportlab.lib import colors
from reportlab.lib.utils import simpleSplit
from typing import List
from .models import CuttingJob
from .utils import (
    plan_cuts_for_job,
    calculate_efficiency,
    PipeAssignment,
    build_results_summary,
)


class CuttingStockUI:
    """
    Main UI controller for the Cutting Stock Optimizer.
    """

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Cutting Stock Optimizer")
        self.root.geometry("1200x800")

        # Track which widget should receive mouse wheel scrolling
        self.active_mousewheel_widget = None

        # Store the latest computed result so it can be exported to PDF
        self.last_assignments = []
        self.last_new_pipe_count = 0
        self.last_efficiency = None
        self.last_kerf = 0
        self.last_summary_text = ""

        # Main tab container
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True, padx=10, pady=10)

        # Input tab
        self.input_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.input_tab, text="Input")

        # Results tab
        self.results_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.results_tab, text="Results")

        self.setup_input_tab()
        self.setup_results_tab()
        self.toggle_leftovers()

        # Global mouse wheel binding
        self.root.bind_all("<MouseWheel>", self._on_mousewheel_windows)
        self.root.bind_all("<Button-4>", self._on_mousewheel_linux)
        self.root.bind_all("<Button-5>", self._on_mousewheel_linux)


    def setup_input_tab(self):
        """
        Build the Input tab with:
        - top buttons
        - settings form
        - two side-by-side panels
        - fixed headers outside the scrollable row areas
        """

        self.input_tab.columnconfigure(0, weight=1)
        self.input_tab.rowconfigure(3, weight=1)

        # ====== TOP BUTTONS ======
        top_button_frame = ttk.Frame(self.input_tab)
        top_button_frame.grid(row=0, column=0, sticky="w", padx=5, pady=5)

        ttk.Button(top_button_frame, text="Save Plan", command=self.save_plan).pack(side="left", padx=(0, 10))
        ttk.Button(top_button_frame, text="Load Plan", command=self.load_plan).pack(side="left", padx=(0, 10))
        ttk.Button(top_button_frame, text="Compute Cutting Plan", command=self.compute_plan).pack(side="left")

        # ====== SETTINGS ======
        form_frame = ttk.Frame(self.input_tab)
        form_frame.grid(row=1, column=0, sticky="w", padx=5, pady=5)

        ttk.Label(form_frame, text="Stock Pipe Length:").grid(
            row=0, column=0, sticky="e", padx=(0, 10), pady=2
        )
        self.stock_len_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.stock_len_var, width=20).grid(
            row=0, column=1, sticky="w", pady=2
        )

        ttk.Label(form_frame, text="Kerf:").grid(
            row=1, column=0, sticky="e", padx=(0, 10), pady=2
        )
        self.kerf_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=self.kerf_var, width=20).grid(
            row=1, column=1, sticky="w", pady=2
        )

        # ====== CHECKBOX ======
        self.include_leftovers_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            self.input_tab,
            text="Include Leftovers",
            variable=self.include_leftovers_var,
            command=self.toggle_leftovers
        ).grid(row=2, column=0, sticky="w", padx=5, pady=5)

        # ====== PANELS AREA ======
        self.panels_frame = ttk.Frame(self.input_tab)
        self.panels_frame.grid(row=3, column=0, sticky="nsw", padx=2, pady=5)

        # Keep panels at natural width instead of stretching them wide
        self.panels_frame.columnconfigure(0, weight=0)
        self.panels_frame.columnconfigure(1, weight=0)
        self.panels_frame.rowconfigure(0, weight=1)

        # =========================
        # REQUIRED CUTS PANEL
        # =========================
        self.cuts_panel = ttk.Frame(self.panels_frame)
        self.cuts_panel.grid(row=0, column=0, sticky="nsw", padx=(0, 10), pady=5)

        self.cuts_panel.columnconfigure(0, weight=1)
        self.cuts_panel.rowconfigure(2, weight=1)

        ttk.Label(
            self.cuts_panel,
            text="Required Cuts",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, pady=(0, 10))

        self.cuts_header_frame = ttk.Frame(self.cuts_panel)
        self.cuts_header_frame.grid(row=1, column=0, sticky="w", pady=(0, 5))

        ttk.Label(self.cuts_header_frame, text="Label", width=18, anchor="center").grid(row=0, column=0, padx=2)
        ttk.Label(self.cuts_header_frame, text="Length", width=12, anchor="center").grid(row=0, column=1, padx=2)
        ttk.Label(self.cuts_header_frame, text="Quantity", width=10, anchor="center").grid(row=0, column=2, padx=2)
        ttk.Label(self.cuts_header_frame, text="", width=10).grid(row=0, column=3, padx=2)

        self.cuts_scroll_container = ttk.Frame(self.cuts_panel, borderwidth=1, relief="solid")
        self.cuts_scroll_container.grid(row=2, column=0, sticky="nsw")

        self.cuts_scroll_container.rowconfigure(0, weight=1)
        self.cuts_scroll_container.columnconfigure(0, weight=1)

        self.cuts_canvas = tk.Canvas(self.cuts_scroll_container, highlightthickness=0, width=460)
        self.cuts_canvas.grid(row=0, column=0, sticky="nsew")

        self.cuts_scrollbar = ttk.Scrollbar(self.cuts_scroll_container, orient="vertical", command=self.cuts_canvas.yview)
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

        # =========================
        # LEFTOVERS PANEL
        # =========================
        self.leftovers_panel = ttk.Frame(self.panels_frame)
        self.leftovers_panel.grid(row=0, column=1, sticky="nsw", padx=(0, 10), pady=5)

        self.leftovers_panel.columnconfigure(0, weight=1)
        self.leftovers_panel.rowconfigure(2, weight=1)

        ttk.Label(
            self.leftovers_panel,
            text="Leftover Pipes",
            font=("Segoe UI", 10, "bold")
        ).grid(row=0, column=0, pady=(0, 10))

        self.leftovers_header_frame = ttk.Frame(self.leftovers_panel)
        self.leftovers_header_frame.grid(row=1, column=0, sticky="w", pady=(0, 5))

        ttk.Label(self.leftovers_header_frame, text="Label", width=18, anchor="center").grid(row=0, column=0, padx=2)
        ttk.Label(self.leftovers_header_frame, text="Length", width=12, anchor="center").grid(row=0, column=1, padx=2)
        ttk.Label(self.leftovers_header_frame, text="Quantity", width=10, anchor="center").grid(row=0, column=2, padx=2)
        ttk.Label(self.leftovers_header_frame, text="", width=10).grid(row=0, column=3, padx=2)

        self.leftovers_scroll_container = ttk.Frame(self.leftovers_panel, borderwidth=1, relief="solid")
        self.leftovers_scroll_container.grid(row=2, column=0, sticky="nsw")

        self.leftovers_scroll_container.rowconfigure(0, weight=1)
        self.leftovers_scroll_container.columnconfigure(0, weight=1)

        self.leftovers_canvas = tk.Canvas(self.leftovers_scroll_container, highlightthickness=0, width=460)
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

    def setup_results_tab(self):
        """
        Build the Results tab with equal-height result areas.
        """
        self.results_tab.columnconfigure(0, weight=1)
        self.results_tab.rowconfigure(0, weight=1)

        # EXPORT PDF BUTTON
        export_button_frame = ttk.Frame(self.results_tab)
        export_button_frame.grid(row=1, column=0, sticky="e", padx=10, pady=(0, 5))

        ttk.Button(
            export_button_frame,
            text="Export as PDF",
            command=self.export_results_pdf
        ).pack(side="right")

        #Output frame to hold both the text and visualization sections, with equal height

        output_frame = ttk.Frame(self.results_tab)
        output_frame.grid(row=0, column=0, sticky="nsew", padx=10, pady=5)

        output_frame.columnconfigure(0, weight=1)
        output_frame.rowconfigure(0, weight=1, uniform="results")
        output_frame.rowconfigure(1, weight=1, uniform="results")

        # ====== RESULTS TEXT SECTION ======
        results_frame = ttk.LabelFrame(output_frame, text="Results", padding=10)
        results_frame.grid(row=0, column=0, sticky="nsew", pady=(0, 5))

        results_frame.rowconfigure(0, weight=1)
        results_frame.columnconfigure(0, weight=1)

        text_frame = ttk.Frame(results_frame)
        text_frame.grid(row=0, column=0, sticky="nsew")

        text_frame.rowconfigure(0, weight=1)
        text_frame.columnconfigure(0, weight=1)

        self.results_text = tk.Text(
            text_frame,
            wrap="word",
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 10),
            height=1
        )
        self.results_text.grid(row=0, column=0, sticky="nsew")

        results_scroll = ttk.Scrollbar(text_frame, orient="vertical", command=self.results_text.yview)
        results_scroll.grid(row=0, column=1, sticky="ns")
        self.results_text.configure(yscrollcommand=results_scroll.set)

        self._bind_mousewheel_target(self.results_text, self.results_text)

        # ====== PIPE VISUALIZATION SECTION ======
        vis_frame = ttk.LabelFrame(output_frame, text="Pipe Visualization", padding=10)
        vis_frame.grid(row=1, column=0, sticky="nsew", pady=(5, 0))

        vis_frame.rowconfigure(0, weight=1)
        vis_frame.columnconfigure(0, weight=1)

        canvas_frame = ttk.Frame(vis_frame)
        canvas_frame.grid(row=0, column=0, sticky="nsew")

        canvas_frame.rowconfigure(0, weight=1)
        canvas_frame.columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(
            canvas_frame,
            bg="white",
            highlightthickness=0,
            height=1
        )
        self.canvas.grid(row=0, column=0, sticky="nsew")

        canvas_vscroll = ttk.Scrollbar(canvas_frame, orient="vertical", command=self.canvas.yview)
        canvas_vscroll.grid(row=0, column=1, sticky="ns")

        canvas_hscroll = ttk.Scrollbar(vis_frame, orient="horizontal", command=self.canvas.xview)
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
        self.cuts_canvas.itemconfig(self.cuts_window, width=event.width)

    def on_leftovers_frame_configure(self, event=None):
        self.leftovers_canvas.configure(scrollregion=self.leftovers_canvas.bbox("all"))

    def on_leftovers_canvas_configure(self, event):
        self.leftovers_canvas.itemconfig(self.leftovers_window, width=event.width)

    def toggle_leftovers(self):
        if self.include_leftovers_var.get():
            self.leftovers_panel.grid()
        else:
            self.leftovers_panel.grid_remove()

    def add_cuts_row(self):
        row_num = len(self.cuts_rows)

        label_entry = ttk.Entry(self.cuts_frame, width=18)
        length_entry = ttk.Entry(self.cuts_frame, width=12)
        qty_entry = ttk.Entry(self.cuts_frame, width=10)

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

        label_entry = ttk.Entry(self.leftovers_frame, width=18)
        length_entry = ttk.Entry(self.leftovers_frame, width=12)
        qty_entry = ttk.Entry(self.leftovers_frame, width=10)

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

    def save_plan(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile="new_cutting_plan.csv"
        )
        if not file_path:
            return

        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["Type", "Label", "Length", "Quantity"])

                writer.writerow(["Settings", "StockPipeLength", self.stock_len_var.get(), ""])
                writer.writerow(["Settings", "Kerf", self.kerf_var.get(), ""])
                writer.writerow(["Settings", "IncludeLeftovers", self.include_leftovers_var.get(), ""])

                cuts = self.parse_grid_input(self.cuts_rows, "cut")
                for label, length, qty in cuts:
                    writer.writerow(["Cut", label, length, qty])

                if self.include_leftovers_var.get():
                    leftovers = self.parse_grid_input(self.leftovers_rows, "Left Over")
                    for label, length, qty in leftovers:
                        writer.writerow(["Leftover", label, length, qty])

            messagebox.showinfo("Success", f"Plan saved to {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save plan: {e}")

    def load_plan(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv")]
        )
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as csvfile:
                reader = csv.reader(csvfile)
                next(reader)

                cuts = []
                leftovers = []
                stock_len = ""
                kerf = ""
                include_leftovers = False

                for row in reader:
                    if len(row) < 4:
                        continue

                    typ, label, length_str, qty_str = row

                    if typ == "Settings":
                        if label == "StockPipeLength":
                            stock_len = length_str
                        elif label == "Kerf":
                            kerf = length_str
                        elif label == "IncludeLeftovers":
                            include_leftovers = length_str.lower() == "true"
                    elif typ == "Cut":
                        cuts.append((label, int(length_str), int(qty_str)))
                    elif typ == "Leftover":
                        leftovers.append((label, int(length_str), int(qty_str)))

                while len(self.cuts_rows) > 0:
                    self.remove_cuts_row(0)
                while len(self.leftovers_rows) > 0:
                    self.remove_leftovers_row(0)

                self.stock_len_var.set(stock_len)
                self.kerf_var.set(kerf)
                self.include_leftovers_var.set(include_leftovers)
                self.toggle_leftovers()

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

            messagebox.showinfo("Success", f"Plan loaded from {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load plan: {e}")

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
                    label = f"{default_prefix} {label_counter:02d}"
                    label_counter += 1

                result.append((label, length, qty))
            except ValueError:
                raise ValueError(f"Invalid length or quantity in row: {label}, {length_str}, {qty_str}")

        return result

    def compute_plan(self):
        try:
            cut_req = self.parse_grid_input(self.cuts_rows, "cut")
            leftovers = {}
            if self.include_leftovers_var.get():
                leftovers = self.parse_grid_input(self.leftovers_rows, "Left Over")

            stock_len_str = self.stock_len_var.get().strip()
            kerf_str = self.kerf_var.get().strip()

            if not stock_len_str or not kerf_str:
                messagebox.showerror("Error", "Stock Pipe Length and Kerf are required")
                return

            try:
                stock_len = int(stock_len_str)
                kerf = int(kerf_str)
            except ValueError:
                messagebox.showerror("Error", "Stock Pipe Length and Kerf must be numbers")
                return

            job = CuttingJob(
                cut_requirements=cut_req,
                stock_pipe_length=stock_len,
                leftover_pipes=leftovers,
                kerf=kerf,
                include_leftovers=self.include_leftovers_var.get(),
            )

            assignments, new_pipe_count = plan_cuts_for_job(job)
            efficiency = calculate_efficiency(assignments)

            # Save latest result for PDF export
            self.last_assignments = assignments
            self.last_new_pipe_count = new_pipe_count
            self.last_efficiency = efficiency
            self.last_kerf = kerf
            self.last_summary_text = build_results_summary(assignments, new_pipe_count, efficiency, kerf)

            # Show text in Results tab
            self.results_text.delete(1.0, tk.END)
            self.results_text.insert(tk.END, self.last_summary_text)

            self.visualize_pipes(assignments, job.kerf)
            self.notebook.select(self.results_tab)

        except Exception as e:
            messagebox.showerror("Error", str(e))

    def visualize_pipes(self, assignments: list[PipeAssignment], kerf: int):
        self.canvas.delete("all")

        y_offset = 20
        max_width = 1600

        max_pipe_length = max((pipe.original_length for pipe in assignments), default=5000)
        scale = max_width / max(5000, max_pipe_length)

        colors = ["red", "blue", "green", "orange", "purple", "brown", "pink", "gray", "cyan", "magenta"]
        font_label = ("Segoe UI", 8, "bold")
        font_pipe = ("Segoe UI", 9, "bold")

        for pipe in assignments:
            if not pipe.cuts:
                continue

            pipe_right = 10 + int(pipe.original_length * scale)
            self.canvas.create_text(
                10, y_offset - 10,
                text=f"Pipe {pipe.id} ({pipe.source})",
                anchor="w",
                font=font_pipe
            )
            self.canvas.create_rectangle(
                10, y_offset, pipe_right, y_offset + 30,
                outline="black", fill="lightgray"
            )

            cut_x = 10.0
            for i, cut in enumerate(pipe.cuts):
                cut_width = cut.length * scale

                if i == len(pipe.cuts) - 1:
                    current_end = cut_x + cut_width
                    if pipe.remaining_length == 0 and current_end < pipe_right:
                        cut_width += (pipe_right - current_end)

                color = colors[i % len(colors)]
                self.canvas.create_rectangle(
                    int(cut_x), y_offset,
                    int(cut_x + cut_width), y_offset + 30,
                    fill=color, outline="black"
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
        """
        EXPORT RESULTS PDF:
        Save the current cutting plan as a PDF with:
        - result summary text at the top
        - visual pipe layout below
        """
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

            # ===== TITLE =====
            pdf.setFont("Helvetica-Bold", 14)
            pdf.drawString(margin, y, "Cutting Stock Optimizer - Cutting Plan")
            y -= 25

            # ===== RESULTS TEXT =====
            pdf.setFont("Helvetica", 10)
            line_height = 14
            max_text_width = page_width - (2 * margin)

            for line in self.last_summary_text.splitlines():
                # Wrap each line to fit the PDF page width
                wrapped_lines = simpleSplit(line, "Helvetica", 10, max_text_width)

                # Keep blank lines
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

            # ===== VISUAL PIPE SECTION =====
            pdf.setFont("Helvetica-Bold", 12)
            pdf.drawString(margin, y, "Pipe Visualization")
            y -= 20

            available_width = page_width - (2 * margin)
            colors_list = [
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

            max_pipe_length = max((pipe.original_length for pipe in self.last_assignments if pipe.cuts), default=5000)
            scale = available_width / max(5000, max_pipe_length)

            bar_height = 18
            pipe_spacing = 38

            for pipe_index, pipe in enumerate(self.last_assignments, start=1):
                if not pipe.cuts:
                    continue

                # Start a new page if needed
                if y < 80:
                    pdf.showPage()
                    y = page_height - margin
                    pdf.setFont("Helvetica-Bold", 12)
                    pdf.setFillColor(colors.black)
                    pdf.drawString(margin, y, "Pipe Visualization (continued)")
                    y -= 20

                # Pipe label
                pdf.setFont("Helvetica-Bold", 10)
                pdf.setFillColor(colors.black)
                pdf.drawString(margin, y, f"Pipe {pipe_index} ({pipe.source})")
                y -= 12

                pipe_x = margin
                pipe_y = y - bar_height
                pipe_width = pipe.original_length * scale

                # Draw full pipe background
                pdf.setStrokeColor(colors.black)
                pdf.setFillColor(colors.lightgrey)
                pdf.rect(pipe_x, pipe_y, pipe_width, bar_height, stroke=1, fill=1)

                # Draw cut sections
                cut_x = pipe_x
                for i, cut in enumerate(pipe.cuts):
                    cut_width = cut.length * scale
                    fill_color = colors_list[i % len(colors_list)]

                    pdf.setFillColor(fill_color)
                    pdf.rect(cut_x, pipe_y, cut_width, bar_height, stroke=1, fill=1)

                    # Draw cut label if there is enough room
                    if cut_width > 45:
                        pdf.setFillColor(colors.white)
                        pdf.setFont("Helvetica-Bold", 8)
                        pdf.drawCentredString(
                            cut_x + (cut_width / 2),
                            pipe_y + 5,
                            f"{cut.id}({cut.length})"
                        )

                    # Move forward, including kerf spacing except after last cut
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

            # Give the app its own Windows identity so the taskbar
            # shows your app icon instead of Python's.
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(
                "Martin.CuttingStockOptimizer.1"
            )
        except Exception:
            pass

    root = tk.Tk()

    # SET APPLICATION ICON
    try:
        ico_path = resource_path(os.path.join("assets", "CSO_icon.ico"))
        png_path = resource_path(os.path.join("assets", "CSO_icon.png"))

        if sys.platform == "win32":
            root.iconbitmap(ico_path)

        icon_image = tk.PhotoImage(file=png_path)
        root.iconphoto(True, icon_image)
        root._icon_image = icon_image  # keep reference alive

    except Exception as e:
        print(f"Warning: Could not load icon: {e}")

    app = CuttingStockUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()