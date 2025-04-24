#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
GUI application to search a lab inventory Excel spreadsheet.
Uses Tkinter for the GUI, pandas for data handling.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, font as tkFont
import pandas as pd
import os
import threading # To run loading/searching in background

# --- Configuration (Defaults) ---
# (These can be adjusted, but the GUI provides controls)
DEFAULT_OUTPUT_COLUMNS = ['Item Name *', 'Location']
FORCE_STRING_COLS = [ # Ensure these are treated as text for searching
    'CAS Number', 'Lot Number', 'Catalog #', 'Serial Number', 'Location',
    'Sub-location', 'Item Name *', 'Genotype *', 'Notes', 'Vendor', 'Alt Name/ID'
]

# --- Core Logic (Slightly modified to return data/errors) ---

def load_inventory(file_path: str, sheet_name=0) -> tuple[pd.DataFrame | None, str | None, list[str] | None, list[str] | None]:
    """
    Loads inventory, returns DataFrame, error message, column names, and sheet names.
    """
    if not file_path or not os.path.exists(file_path):
        return None, f"Error: File not found or path is invalid: '{file_path}'", None, None

    all_sheet_names = []
    try:
        # Get sheet names first without loading all data initially
        xls = pd.ExcelFile(file_path)
        all_sheet_names = xls.sheet_names

        # Determine the sheet to load
        actual_sheet_name = None
        if isinstance(sheet_name, int):
            if 0 <= sheet_name < len(all_sheet_names):
                actual_sheet_name = all_sheet_names[sheet_name]
            else:
                 return None, f"Error: Sheet index {sheet_name} is out of range.", None, all_sheet_names
        elif isinstance(sheet_name, str):
            if sheet_name in all_sheet_names:
                actual_sheet_name = sheet_name
            else:
                return None, f"Error: Sheet name '{sheet_name}' not found.", None, all_sheet_names
        else:
             return None, "Error: Invalid sheet identifier type.", None, all_sheet_names

        # Now load the specific sheet
        df = pd.read_excel(file_path, sheet_name=actual_sheet_name)

        # Basic cleaning
        for col in FORCE_STRING_COLS:
            if col in df.columns:
                try:
                    df[col] = df[col].astype("string").fillna('')
                except (TypeError, AttributeError):
                    df[col] = df[col].astype(str).replace('nan', '').replace('<NA>', '').fillna('')

        return df, None, df.columns.tolist(), all_sheet_names # Success: DataFrame, no error, columns, sheets

    except FileNotFoundError:
        return None, f"Error: File not found at '{file_path}'", None, None
    except ImportError:
         return None, "Error: Missing pandas or openpyxl. Run: pip install pandas openpyxl", None, None
    except Exception as e:
        return None, f"An error occurred loading '{os.path.basename(file_path)}':\n{e}", None, all_sheet_names


def search_inventory(df: pd.DataFrame, query: str, searchable_columns: list[str]) -> tuple[pd.DataFrame | None, str | None]:
    """
    Searches DataFrame, returns results DataFrame and error message.
    """
    if df is None or df.empty:
        return pd.DataFrame(), "Info: Inventory data is not loaded or is empty." # Return empty df, info message

    if not query.strip():
        return pd.DataFrame(), "Info: Please enter a search query." # Return empty df, info message

    search_terms = query.lower().split()

    valid_searchable_cols = [col for col in searchable_columns if col in df.columns]
    if not valid_searchable_cols:
         return pd.DataFrame(), "Error: No valid columns selected or available for searching."

    df_search = df[valid_searchable_cols].copy()
    for col in valid_searchable_cols:
        df_search[col] = df_search[col].astype(str).fillna('')

    combined_mask = pd.Series([True] * len(df), index=df.index)

    try:
        for term in search_terms:
            term_mask_for_row = df_search.apply(
                lambda row: row.str.contains(term, case=False, na=False, regex=False).any(),
                axis=1
            )
            combined_mask &= term_mask_for_row

        results = df.loc[combined_mask]
        return results, None # Success: results DataFrame, no error

    except Exception as e:
        return None, f"Error during search: {e}"


# --- GUI Application Class ---

class InventorySearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lab Inventory Search")
        self.root.geometry("900x700") # Adjust size as needed

        # Style configuration
        style = ttk.Style(self.root)
        style.theme_use('clam') # Or 'alt', 'default', 'vista', etc.

        # Data storage
        self.inventory_df = None
        self.file_path = tk.StringVar()
        self.sheet_names = []
        self.selected_sheet = tk.StringVar()
        self.column_names = []
        self.search_vars = {} # {col_name: tk.BooleanVar()}
        self.display_vars = {} # {col_name: tk.BooleanVar()}
        self.show_all_cols_var = tk.BooleanVar(value=False)

        # --- Top Frame: File and Sheet Selection ---
        top_frame = ttk.Frame(root, padding="10")
        top_frame.pack(fill=tk.X, side=tk.TOP)

        ttk.Button(top_frame, text="Select Inventory File (.xlsx)", command=self.select_file).pack(side=tk.LEFT, padx=5)
        ttk.Entry(top_frame, textvariable=self.file_path, state='readonly', width=50).pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)

        ttk.Label(top_frame, text="Sheet:").pack(side=tk.LEFT, padx=(10, 0))
        self.sheet_combo = ttk.Combobox(top_frame, textvariable=self.selected_sheet, state='readonly', width=20)
        self.sheet_combo.pack(side=tk.LEFT, padx=5)
        self.sheet_combo.bind('<<ComboboxSelected>>', self.on_sheet_change) # Reload if sheet changes

        # --- Middle Frame: Search Query and Controls ---
        middle_frame = ttk.Frame(root, padding="10")
        middle_frame.pack(fill=tk.X, side=tk.TOP)

        ttk.Label(middle_frame, text="Search Query:").pack(side=tk.LEFT, padx=5)
        self.query_entry = ttk.Entry(middle_frame, width=40)
        self.query_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5)
        self.query_entry.bind("<Return>", self.run_search) # Allow pressing Enter to search

        self.search_button = ttk.Button(middle_frame, text="Search", command=self.run_search, state=tk.DISABLED)
        self.search_button.pack(side=tk.LEFT, padx=5)

        # --- Checkbox Frames Area ---
        check_area_frame = ttk.Frame(root, padding="5")
        check_area_frame.pack(fill=tk.X, side=tk.TOP)

        # Search Columns Frame
        search_cols_frame_outer = ttk.LabelFrame(check_area_frame, text="Search In Columns", padding="5")
        search_cols_frame_outer.pack(side=tk.LEFT, fill=tk.Y, padx=5, expand=True)
        search_cols_canvas = tk.Canvas(search_cols_frame_outer, borderwidth=0, width=180) # Fixed width for check area
        search_cols_scrollbar = ttk.Scrollbar(search_cols_frame_outer, orient="vertical", command=search_cols_canvas.yview)
        self.search_cols_frame = ttk.Frame(search_cols_canvas) # Frame inside canvas
        self.search_cols_frame.bind("<Configure>", lambda e: search_cols_canvas.configure(scrollregion=search_cols_canvas.bbox("all")))
        search_cols_canvas.create_window((0, 0), window=self.search_cols_frame, anchor="nw")
        search_cols_canvas.configure(yscrollcommand=search_cols_scrollbar.set)
        search_cols_canvas.pack(side="left", fill="both", expand=True)
        search_cols_scrollbar.pack(side="right", fill="y")
        ttk.Button(search_cols_frame_outer, text="All/None", command=lambda: self.toggle_all_checks(self.search_vars)).pack(side=tk.BOTTOM, fill=tk.X, pady=(5,0))


        # Display Columns Frame
        display_cols_frame_outer = ttk.LabelFrame(check_area_frame, text="Display Columns", padding="5")
        display_cols_frame_outer.pack(side=tk.LEFT, fill=tk.Y, padx=5, expand=True)
        display_cols_canvas = tk.Canvas(display_cols_frame_outer, borderwidth=0, width=180) # Fixed width
        display_cols_scrollbar = ttk.Scrollbar(display_cols_frame_outer, orient="vertical", command=display_cols_canvas.yview)
        self.display_cols_frame = ttk.Frame(display_cols_canvas) # Frame inside canvas
        self.display_cols_frame.bind("<Configure>", lambda e: display_cols_canvas.configure(scrollregion=display_cols_canvas.bbox("all")))
        display_cols_canvas.create_window((0, 0), window=self.display_cols_frame, anchor="nw")
        display_cols_canvas.configure(yscrollcommand=display_cols_scrollbar.set)
        display_cols_canvas.pack(side="left", fill="both", expand=True)
        display_cols_scrollbar.pack(side="right", fill="y")
        control_frame = ttk.Frame(display_cols_frame_outer) # Frame for buttons below checkboxes
        control_frame.pack(side=tk.BOTTOM, fill=tk.X, pady=(5,0))
        ttk.Checkbutton(control_frame, text="Show All Columns", variable=self.show_all_cols_var, command=self.toggle_display_cols_enable).pack(side=tk.LEFT, padx=2)
        ttk.Button(control_frame, text="All/None", command=lambda: self.toggle_all_checks(self.display_vars, respect_show_all=True)).pack(side=tk.RIGHT, padx=2)

        # --- Results Frame ---
        results_frame = ttk.Frame(root, padding="10")
        results_frame.pack(expand=True, fill=tk.BOTH, side=tk.TOP)

        # Treeview for results table
        self.tree = ttk.Treeview(results_frame, show='headings') # show='headings' hides the default first empty column
        vsb = ttk.Scrollbar(results_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(results_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(expand=True, fill=tk.BOTH)

        # --- Status Bar ---
        self.status_var = tk.StringVar()
        status_bar = ttk.Label(root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W, padding="2 5")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self.set_status("Please select an inventory file.")

    def set_status(self, message):
        self.status_var.set(message)
        self.root.update_idletasks() # Force GUI update

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Select Excel Inventory File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if path:
            self.file_path.set(path)
            self.load_file_data(path) # Load initial sheet

    def load_file_data(self, path, sheet_identifier=0):
        self.inventory_df = None # Clear previous data
        self.column_names = []
        self.sheet_names = []
        self.update_checkboxes() # Clear checkboxes first
        self.clear_results_table()
        self.search_button.config(state=tk.DISABLED)
        self.sheet_combo['values'] = []
        self.selected_sheet.set('')

        self.set_status(f"Loading '{os.path.basename(path)}'...")

        # Run loading in a separate thread to avoid freezing GUI
        thread = threading.Thread(target=self._load_worker, args=(path, sheet_identifier), daemon=True)
        thread.start()

    def _load_worker(self, path, sheet_identifier):
        df, error_msg, cols, sheets = load_inventory(path, sheet_identifier)

        # Use schedule to update GUI from the main thread
        self.root.after(0, self._load_complete, df, error_msg, cols, sheets)

    def _load_complete(self, df, error_msg, cols, sheets):
        if error_msg:
            self.set_status(error_msg)
            messagebox.showerror("Loading Error", error_msg)
            self.file_path.set("") # Clear invalid file path
            if sheets: # Still populate sheets if loading failed but we got the names
                self.sheet_names = sheets
                self.sheet_combo['values'] = self.sheet_names
        elif df is not None:
            self.inventory_df = df
            self.column_names = cols if cols else []
            self.sheet_names = sheets if sheets else []

            # Update Sheet Combobox
            self.sheet_combo['values'] = self.sheet_names
            # Try to find the loaded sheet name and set it
            loaded_sheet_name = self.inventory_df.attrs.get('sheet_name', self.sheet_names[0] if self.sheet_names else '') # Hacky way, need better sheet tracking in load_inventory if possible
            if self.selected_sheet.get() not in self.sheet_names: # if current selection is invalid or empty
                 try:
                      # Try to infer loaded sheet if index was used
                      if isinstance(self.sheet_combo.current(), int) and self.sheet_combo.current() >= 0:
                           self.selected_sheet.set(self.sheet_names[self.sheet_combo.current()])
                      else: # Default to first sheet
                           self.selected_sheet.set(self.sheet_names[0])
                 except IndexError:
                     self.selected_sheet.set("") # Should not happen if sheets is populated


            self.update_checkboxes()
            self.search_button.config(state=tk.NORMAL)
            self.set_status(f"Loaded {len(self.inventory_df):,} items from sheet '{self.selected_sheet.get()}'. Ready to search.")
        else:
             self.set_status("Loading failed. Unknown error.") # Should have error_msg

    def on_sheet_change(self, event=None):
        selected = self.selected_sheet.get()
        if selected and self.file_path.get():
            # Reload data for the newly selected sheet
            self.load_file_data(self.file_path.get(), selected)

    def update_checkboxes(self):
        # Clear existing checkboxes first
        for widget in self.search_cols_frame.winfo_children():
            widget.destroy()
        for widget in self.display_cols_frame.winfo_children():
            widget.destroy()
        self.search_vars.clear()
        self.display_vars.clear()

        # Create new checkboxes based on self.column_names
        for col in self.column_names:
            # Search Column Checkboxes
            s_var = tk.BooleanVar()
            # Default check common columns like 'Item Name *' if they exist
            if col in ['Item Name *', 'Catalog #', 'Location', 'Notes', 'CAS Number']:
                 s_var.set(True)
            cb_s = ttk.Checkbutton(self.search_cols_frame, text=col, variable=s_var)
            cb_s.pack(anchor=tk.W, padx=5, pady=1)
            self.search_vars[col] = s_var

            # Display Column Checkboxes
            d_var = tk.BooleanVar()
            if col in DEFAULT_OUTPUT_COLUMNS: # Default check specified output columns
                d_var.set(True)
            cb_d = ttk.Checkbutton(self.display_cols_frame, text=col, variable=d_var)
            cb_d.pack(anchor=tk.W, padx=5, pady=1)
            self.display_vars[col] = d_var

        # Ensure display checkboxes are enabled/disabled correctly based on "Show All"
        self.toggle_display_cols_enable()


    def toggle_all_checks(self, var_dict, respect_show_all=False):
        # Determine the target state (toggle based on the first item's current state)
        if not var_dict: return
        current_state = list(var_dict.values())[0].get()
        new_state = not current_state

        # If toggling display vars and "Show All" is checked, do nothing
        if respect_show_all and self.show_all_cols_var.get():
             return

        for var in var_dict.values():
            var.set(new_state)

    def toggle_display_cols_enable(self):
        # Disable individual display checkboxes if "Show All" is checked
        state = tk.DISABLED if self.show_all_cols_var.get() else tk.NORMAL
        for widget in self.display_cols_frame.winfo_children():
            if isinstance(widget, ttk.Checkbutton):
                widget.config(state=state)

    def run_search(self, event=None): # event=None allows binding to Enter key
        if not self.inventory_df is not None:
             self.set_status("Error: Inventory not loaded.")
             messagebox.showerror("Error", "Please load an inventory file first.")
             return

        query = self.query_entry.get()
        search_cols = [col for col, var in self.search_vars.items() if var.get()]

        if not search_cols:
            self.set_status("Error: Please select at least one column to search in.")
            messagebox.showwarning("Input Error", "Please select at least one column under 'Search In Columns'.")
            return

        self.set_status(f"Searching for '{query}'...")
        self.clear_results_table()

        # Run search in a thread
        thread = threading.Thread(target=self._search_worker, args=(self.inventory_df, query, search_cols), daemon=True)
        thread.start()

    def _search_worker(self, df, query, search_cols):
        results_df, error_msg = search_inventory(df, query, search_cols)
        # Schedule GUI update from main thread
        self.root.after(0, self._search_complete, results_df, error_msg)

    def _search_complete(self, results_df, error_msg):
        if error_msg:
            self.set_status(error_msg)
            if "Error:" in error_msg:
                 messagebox.showerror("Search Error", error_msg)
            # Don't show messagebox for info messages like empty query
            return

        if results_df.empty:
            self.set_status(f"No results found for '{self.query_entry.get()}'.")
            return

        # Determine columns to display
        if self.show_all_cols_var.get():
            display_cols = self.column_names
        else:
            display_cols = [col for col, var in self.display_vars.items() if var.get()]
            if not display_cols: # Fallback if user somehow deselected all
                display_cols = DEFAULT_OUTPUT_COLUMNS
                # Ensure fallback columns exist
                display_cols = [col for col in display_cols if col in results_df.columns]
                if not display_cols and not self.show_all_cols_var.get(): # Absolute fallback
                     display_cols = results_df.columns.tolist()[:5] # Show first few if defaults are gone

        # Filter results DataFrame to only include display columns that actually exist
        final_display_cols = [col for col in display_cols if col in results_df.columns]
        results_to_display = results_df[final_display_cols]


        # Update Treeview
        self.tree['columns'] = final_display_cols
        self.tree['displaycolumns'] = final_display_cols # Necessary if columns have spaces

        # Get system font for sensible column widths
        default_font = tkFont.nametofont("TkDefaultFont")

        for col in final_display_cols:
            self.tree.heading(col, text=col)
            # Estimate width based on header and maybe some sample data (can be slow)
            # Simple estimation based on header length for now
            col_width = default_font.measure(col) + 20 # Measure header width + padding
            # Or set a reasonable default/min width
            min_width = 80
            self.tree.column(col, anchor=tk.W, width=max(col_width, min_width), stretch=True)


        # Insert data rows
        for index, row in results_to_display.iterrows():
            # Convert row values to strings for display, handle NaN/None
            values = [str(row[col]) if pd.notna(row[col]) else "" for col in final_display_cols]
            self.tree.insert('', tk.END, values=values)

        self.set_status(f"Found {len(results_df):,} matching items.")


    def clear_results_table(self):
        # Delete all items from the treeview
        for i in self.tree.get_children():
            self.tree.delete(i)
        # Reset columns/headings (optional, can leave them)
        # self.tree['columns'] = []
        # self.tree['displaycolumns'] = []

# --- Main Execution ---
if __name__ == "__main__":
    root = tk.Tk()
    app = InventorySearchApp(root)
    root.mainloop()