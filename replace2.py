import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import logging

# Configure logging
logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(levelname)s - %(message)s',
                    handlers=[
                        logging.FileHandler("replace2.log"),
                        logging.StreamHandler()
                    ])

class ReplacementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Replacement Tool")
        self.root.geometry("600x400")

        # File Path Selection
        self.frame_file = ttk.LabelFrame(root, text="Excel File Selection", padding="5")
        self.frame_file.pack(fill="x", padx=5, pady=5)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(self.frame_file, textvariable=self.path_var, width=50)
        self.path_entry.pack(side="left", padx=5)

        self.browse_btn = ttk.Button(self.frame_file, text="Browse", command=self.browse_file)
        self.browse_btn.pack(side="left", padx=5)

        # Sheet Selection
        self.frame_sheet = ttk.LabelFrame(root, text="Sheet Selection", padding="5")
        self.frame_sheet.pack(fill="x", padx=5, pady=5)

        self.sheet_var = tk.StringVar()
        self.sheet_combo = ttk.Combobox(self.frame_sheet, textvariable=self.sheet_var, state="readonly")
        self.sheet_combo.pack(fill="x", padx=5)

        # Range Input
        self.frame_range = ttk.LabelFrame(root, text="Range Selection (e.g., E2:F26)", padding="5")
        self.frame_range.pack(fill="x", padx=5, pady=5)

        self.range_var = tk.StringVar()
        self.range_entry = ttk.Entry(self.frame_range, textvariable=self.range_var)
        self.range_entry.pack(fill="x", padx=5)

        # Replacement Rules
        self.frame_rules = ttk.LabelFrame(root, text="Replacement Rules (format: find1,replace1,find2,replace2,...)", padding="5")
        self.frame_rules.pack(fill="x", padx=5, pady=5)

        self.rules_text = tk.Text(self.frame_rules, height=5)
        self.rules_text.pack(fill="x", padx=5)

        # Process Button
        self.process_btn = ttk.Button(root, text="Process Replacements", command=self.process_replacements)
        self.process_btn.pack(pady=20)

    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if filename:
            self.path_var.set(filename)
            self.update_sheet_list(filename)

    def update_sheet_list(self, filename):
        try:
            logging.info(f"Reading Excel file: {filename}")
            excel_file = pd.ExcelFile(filename)
            sheet_names = excel_file.sheet_names
            self.sheet_combo['values'] = sheet_names
            if sheet_names:
                self.sheet_combo.set(sheet_names[0])
            logging.info(f"Available sheets: {sheet_names}")
        except Exception as e:
            logging.error(f"Error reading Excel file: {e}")
            messagebox.showerror("Error", f"Error reading Excel file: {str(e)}")

    def safe_replace(self, value, replace_dict):
        if pd.isna(value):
            return value
        try:
            str_value = str(value)  # Convert value to string without normalizing
            for find_str, replace_str in replace_dict.items():
                if find_str in str_value:
                    logging.info(f"Replacing '{find_str}' with '{replace_str}' in '{str_value}'")
                    str_value = str_value.replace(find_str, replace_str)
            return str_value
        except Exception as e:
            logging.error(f"Error replacing value: {e}")
            return value

    def process_replacements(self):
        try:
            # Get input values
            file_path = self.path_var.get()
            sheet_name = self.sheet_var.get()
            cell_range = self.range_var.get()
            rules_text = self.rules_text.get("1.0", "end-1c")

            if not all([file_path, sheet_name, cell_range, rules_text]):
                logging.error("Missing input fields")
                messagebox.showerror("Error", "Please fill in all fields")
                return

            # Parse replacement rules
            rules = rules_text.strip().split(',')
            if len(rules) % 2 != 0:
                logging.error("Invalid replacement rules format")
                messagebox.showerror("Error", "Invalid replacement rules format")
                return

            # Create replacement dictionary
            replace_dict = {rules[i]: rules[i+1] for i in range(0, len(rules), 2)}
            logging.info(f"Replacement rules: {replace_dict}")

            # Read all sheets from the Excel file
            excel_file = pd.ExcelFile(file_path)
            all_sheets = {}
            for sheet in excel_file.sheet_names:
                all_sheets[sheet] = pd.read_excel(file_path, sheet_name=sheet)

            # Modify only the selected sheet and range
            df = all_sheets[sheet_name]
            start_cell, end_cell = cell_range.split(':')
            start_row = int(start_cell[1:]) - 1
            end_row = int(end_cell[1:])
            start_col = ord(start_cell[0].upper()) - ord('A')
            end_col = ord(end_cell[0].upper()) - ord('A') + 1

            # Ensure the range is within bounds
            max_row, max_col = df.shape
            if end_row > max_row:
                end_row = max_row
            if end_col > max_col:
                end_col = max_col

            # Apply replacements only to string columns in the specified range
            for i in range(start_row, end_row):
                for j in range(start_col, end_col):
                    if i < max_row and j < max_col:
                        cell_value = df.iat[i, j]
                        df.iat[i, j] = self.safe_replace(cell_value, replace_dict)

            all_sheets[sheet_name] = df

            # Save all sheets back to the original file
            logging.info(f"Attempting to save back to the file: {file_path}")
            with pd.ExcelWriter(file_path, mode='w') as writer:
                for sheet, data in all_sheets.items():
                    data.to_excel(writer, sheet_name=sheet, index=False)

            logging.info("Replacements completed and saved to the original file")
            messagebox.showinfo("Success", "Replacements completed and saved to the original file!")

        except PermissionError as e:
            logging.error(f"Permission denied: {e}")
            messagebox.showerror("Error", f"Permission denied: {str(e)}")
        except Exception as e:
            logging.error(f"An error occurred: {e}")
            messagebox.showerror("Error", f"An error occurred: {str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = ReplacementApp(root)
    root.mainloop()