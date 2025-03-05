import os
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
from datetime import datetime, timedelta
import subprocess
import openpyxl

# Global employee list
employee_list = {}
history_window_open = False
history_window = None
current_employee_file = None
current_dat_file = None


# Global variables
history_window_open = False
history_window = None  # Ensure global reference to history window


def create_database():
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()

    # Create conversion_history table
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS conversions (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        converted_at TEXT,
        output_path TEXT
    )
    """)

    # Create employees table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS employees (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL
        )
    """)

    conn.commit()
    conn.close()

def save_to_database(filename, output_path):
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO conversions (filename, converted_at, output_path) VALUES (?, ?, ?)", 
                   (filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), output_path))
    conn.commit()
    conn.close()

def filter_in_out_entries(df):
    if df.shape[1] < 2:
        return df  # Return original if insufficient columns

    first_col = df.columns[0]  # Employee Name column
    time_col = df.columns[1]  # Timestamp column

    df[time_col] = pd.to_datetime(df[time_col], errors='coerce')
    df['Date'] = df[time_col].dt.strftime('%Y-%m-%d')  # Extract date
    df['Time'] = df[time_col].dt.strftime('%H:%M:%S')  # Extract time only

    grouped = df.groupby([first_col, 'Date'])['Time'].agg(list).reset_index()
    grouped['Time In'] = grouped['Time'].apply(lambda x: x[0])  # First time log
    grouped['Time Out'] = grouped['Time'].apply(lambda x: x[-1] if len(x) > 1 else "No Out")  # Check for single entry
    grouped = grouped.drop(columns=['Time'])


    if not df.empty:
        last_recorded_date = df[time_col].max().strftime('%Y-%m-%d')  # Get last date in .dat file
        year, month = df[time_col].min().year, df[time_col].min().month
        first_day = datetime(year, month, 1)
        last_day = (first_day.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
        all_dates = pd.date_range(first_day, last_day).strftime('%Y-%m-%d')

        all_employees = grouped['Name'].unique()
        full_index = pd.MultiIndex.from_product([all_employees, all_dates], names=['Name', 'Date'])
        grouped = grouped.set_index(['Name', 'Date']).reindex(full_index).reset_index()

    # Determine day of the week for each date
    grouped['DayOfWeek'] = grouped['Date'].apply(lambda x: datetime.strptime(x, '%Y-%m-%d').strftime('%A'))

    # Fill missing Time In and Time Out
    def mark_absences(row):
        if pd.isna(row['Time In']) and pd.isna(row['Time Out']):
            if row['DayOfWeek'] == 'Saturday':
                return "Saturday", "Saturday"
            elif row['DayOfWeek'] == 'Sunday':
                return "Sunday", "Sunday"
            elif row['Date'] <= last_recorded_date:
                return "Absent", "Absent"
        return row['Time In'], row['Time Out']

    grouped[['Time In', 'Time Out']] = grouped.apply(mark_absences, axis=1, result_type="expand")

    # Convert NaN to empty strings for dates after the last recorded date
    grouped.fillna('', inplace=True)

    # Add Employee No. based on the name mapping
    name_to_id = {v: k for k, v in employee_list.items()}
    grouped['Employee No.'] = grouped['Name'].map(name_to_id)

    result = grouped.pivot(index=['Employee No.', 'Name'], columns='Date', values=['Time In', 'Time Out'])
    result = result.swaplevel(axis=1).sort_index(axis=1, level=0)
    result.columns = [f"{date} ({datetime.strptime(date, '%Y-%m-%d').strftime('%A')}) {status}" for date, status in result.columns]


    return result.reset_index().sort_values(by="Employee No.").reset_index(drop=True)

import openpyxl

def convert_batch_to_excel(files):
    for dat_file in files:
        try:
            df = pd.read_csv(dat_file, delimiter="\t", header=None)

            if df.empty:
                messagebox.showerror("Error", f"The file {dat_file} is empty.")
                return

            if df.shape[1] < 2:
                messagebox.showerror("Error", "DAT file must have at least two columns (ID and Timestamp).")
                return

            df.columns = ["Name", "Timestamp"] + [f"Col_{i}" for i in range(2, df.shape[1])]
            df["Name"] = df["Name"].map(lambda x: employee_list.get(int(x), str(x)))

            df["Timestamp"] = pd.to_datetime(df["Timestamp"], errors="coerce")
            df = df.dropna(subset=["Timestamp"])  # Remove invalid timestamps

            if df.empty:
                messagebox.showerror("Error", "No valid timestamps found in the DAT file.")
                return

            # Extract the month and year for sheet naming and headers
            first_date = df["Timestamp"].min()
            month_year = first_date.strftime("%B %Y")  # Example: "January 2025"
            sheet_name = f"DTR - {month_year}"

            # Filter for first time-in and last time-out
            final_df = filter_in_out_entries(df)

            if final_df.empty:
                messagebox.showerror("Error", "No data available for conversion. Check the input file.")
                return

            # Get save path from user
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=os.path.basename(dat_file).replace(".dat", ".xlsx"),
                title="Save Converted Excel File"
            )

            if save_path:
                with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                    # Write the attendance sheet with a month-based name
                    final_df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)

                    # Auto-adjust column widths
                    writer.book[sheet_name]["A1"] = month_year  # Set month-year as title
                    writer.book[sheet_name]["A1"].font = Font(size=14, bold=True)
                    writer.book[sheet_name]["A1"].alignment = Alignment(horizontal="center")
                    writer.book[sheet_name].merge_cells('A1:E1')  # Merge across columns for better visibility

                    # Generate individual employee DTR sheets
                    for name in df["Name"].unique():
                        generate_employee_dtr(writer, df, name)

                # Auto-adjust column widths
                auto_adjust_column_widths(save_path)
                
                save_to_database(os.path.basename(dat_file), save_path)

                open_file = messagebox.askyesno("Conversion Complete", "File converted successfully!\nDo you want to open it now?")
                if open_file:
                    subprocess.run(["start", "", save_path], shell=True)

                # Remove the .dat file from the listbox after successful conversion
                employee_list_entry.delete(employee_list_entry.get(0, tk.END).index(dat_file))  # Remove the file from the listbox

        except Exception as e:
            messagebox.showerror
        except Exception as e:
            messagebox.showerror

def upload_employee_list():
    global employee_list
    file_path = filedialog.askopenfilename(
        title="Select Employee List TXT File",
        filetypes=[("Text Files", "*.txt")]
    )

    if file_path:
        try:
            conn = sqlite3.connect("conversion_history.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM employees")  # Clear existing list

            with open(file_path, 'r') as file:
                employee_list.clear()
                for line in file:
                    parts = line.strip().split(' ', 1)
                    if len(parts) == 2:
                        emp_id = int(parts[0])
                        emp_name = parts[1].strip().upper()
                        employee_list[emp_id] = emp_name
                        cursor.execute("INSERT INTO employees (id, name) VALUES (?, ?)", (emp_id, emp_name))

            conn.commit()
            conn.close()
            employee_list_entry.insert(tk.END, file_path)  # Show file name in the listbox
            messagebox.showinfo("Success", "Employee list saved and updated successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to upload employee list: {e}")

def load_employee_list():
    global employee_list
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("SELECT id, name FROM employees")
    rows = cursor.fetchall()
    employee_list = {emp_id: name for emp_id, name in rows}
    conn.close()


def generate_employee_dtr(writer, df, employee_name):
    """Generates a formatted Daily Time Record (DTR) sheet for an employee with full month dates."""
    employee_df = df[df["Name"] == employee_name].copy()
    if employee_df.empty:
        return
    
    # Generate all dates for the employee's month
    first_date = employee_df["Timestamp"].min().date()
    year, month = first_date.year, first_date.month
    first_day = datetime(year, month, 1)
    last_day = (first_day.replace(day=28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)
    all_dates = pd.date_range(first_day, last_day).date

    # Prepare full date range DataFrame
    date_df = pd.DataFrame({"Date": all_dates})
    date_df["Weekday"] = date_df["Date"].apply(lambda x: x.strftime('%A'))

    # Extract time in and out per day
    employee_df["Date"] = employee_df["Timestamp"].dt.date
    grouped = employee_df.groupby("Date")["Time"].agg(list).reset_index()
    grouped["Time In"] = grouped["Time"].apply(lambda x: x[0] if len(x) > 0 else "No In")
    grouped["Time Out"] = grouped["Time"].apply(lambda x: x[-1] if len(x) > 1 else "No Out")
    grouped.drop(columns=["Time"], inplace=True)

    # Merge full date range with the employee data
    final_df = pd.merge(date_df, grouped, on="Date", how="left")

    # Fill missing Time In and Time Out
    def fill_missing(row):
        if pd.isnull(row["Time In"]):
            if row["Weekday"] == "Saturday":
                row["Time In"] = "Saturday"
                row["Time Out"] = "Saturday"
            elif row["Weekday"] == "Sunday":
                row["Time In"] = "Sunday"
                row["Time Out"] = "Sunday"
            else:
                row["Time In"] = "Absent"
                row["Time Out"] = "Absent"
        return row

    final_df = final_df.apply(fill_missing, axis=1)

    # Count Saturdays with Time In
    saturday_logs = final_df[(final_df["Weekday"] == "Saturday") & (final_df["Time In"] != "Saturday")].shape[0]
    final_df.loc[0, "Saturdays"] = saturday_logs

    # Format column order
    final_df = final_df[["Date", "Weekday", "Time In", "Time Out"]]

    # Write employee sheet
    final_df.to_excel(writer, sheet_name=employee_name, index=False, startrow=7)

    # Auto-adjust column widths
    worksheet = writer.book[employee_name]
    worksheet["A1"] = employee_name.upper()
    worksheet["A1"].font = Font(size=14, bold=True)
    worksheet["A1"].alignment = Alignment(horizontal="center")
    worksheet.merge_cells("A1:D1")

    worksheet["A3"] = f"For the month of {first_day.strftime('%B %d')} - {last_day.strftime('%d, %Y')}"
    worksheet["A4"] = "Official hours for arrival and departure"
    worksheet["A5"] = "Regular days: 7:00 AM - 4:00 PM"
    worksheet["A6"] = f"Saturdays: {saturday_logs}".ljust(20,)

    for col in worksheet.iter_cols(min_row=8, max_row=worksheet.max_row):
        max_length = max(len(str(cell.value)) if cell.value else 0 for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max_length + 2
        for cell in col:
            cell.alignment = Alignment(horizontal="center")



from openpyxl.utils import get_column_letter
from openpyxl import load_workbook

def auto_adjust_column_widths(file_path):
    wb = load_workbook(file_path)
    
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        column_widths = {}

        for row in ws.iter_rows():
            for cell in row:
                if cell.value:
                    # Check if the cell is part of a merged range
                    if any(cell.coordinate in merged_range for merged_range in ws.merged_cells):
                        continue  # Skip merged cells
                    
                    col_letter = get_column_letter(cell.column)
                    column_widths[col_letter] = max(column_widths.get(col_letter, 0), len(str(cell.value)))

        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width + 2  # Add padding

    wb.save(file_path)

def show_history():
    global history_window, history_window_open

    if not history_window_open:
        history_window_open = True

        def filter_history():
            search_query = search_entry.get().lower()
            for item in tree.get_children():
                tree.delete(item)
            conn = sqlite3.connect("conversion_history.db")
            cursor = conn.cursor()
            cursor.execute("SELECT filename, converted_at, output_path FROM conversions ORDER BY id DESC")
            records = cursor.fetchall()
            conn.close()
            for record in records:
                if search_query in record[0].lower() or search_query in record[1]:
                    tree.insert("", "end", values=record)

        history_window = tk.Toplevel(root)
        history_window.title("Conversion History")
        history_window.geometry("550x350")
        history_window.configure(bg="#f5f5f5")

        search_frame = tk.Frame(history_window, bg="#f5f5f5")
        search_frame.pack(pady=5)
        search_label = tk.Label(search_frame, text="Search:", bg="#f5f5f5")
        search_label.pack(side=tk.LEFT, padx=5)
        search_entry = tk.Entry(search_frame)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_button = tk.Button(search_frame, text="Filter", command=filter_history)
        search_button.pack(side=tk.LEFT, padx=5)

        tree = ttk.Treeview(history_window, columns=("Filename", "Date", "Output Path"), show="headings")
        tree.heading("Filename", text="Filename")
        tree.heading("Date", text="Date Converted")
        tree.heading("Output Path", text="Output Path")
        tree.column("Filename", width=150)
        tree.column("Date", width=120)
        tree.column("Output Path", width=200)
        tree.pack(expand=True, fill="both")
        filter_history()

        history_window.protocol("WM_DELETE_WINDOW", close_history_window)
    else:
        messagebox.showinfo("Info", "History window is already open.")

def close_history_window():
    global history_window, history_window_open
    if history_window:
        history_window.destroy()
        history_window = None
    history_window_open = False

def browse_dat_files():
    file_paths = filedialog.askopenfilenames(filetypes=[("Data Files", "*.dat")])
    if not file_paths:
        return

    for file_path in file_paths:
        employee_list_entry.insert(tk.END, file_path)  # Add filename to the DAT listbox

def upload_employee_list():
    global employee_list
    file_path = filedialog.askopenfilename(
        title="Select Employee List TXT File",
        filetypes=[("Text Files", "*.txt")]
    )

    if file_path:
        try:
            conn = sqlite3.connect("conversion_history.db")
            cursor = conn.cursor()
            cursor.execute("DELETE FROM employees")  # Clear existing list

            with open(file_path, 'r') as file:
                employee_list.clear()
                for line in file:
                    parts = line.strip().split(' ', 1)
                    if len(parts) == 2:
                        emp_id = int(parts[0])
                        emp_name = parts[1].strip().upper()
                        employee_list[emp_id] = emp_name
                        cursor.execute("INSERT INTO employees (id, name) VALUES (?, ?)", (emp_id, emp_name))

            conn.commit()
            conn.close()
            employee_list_entry.insert(tk.END, file_path)  # Show file name in the TXT listbox
            messagebox.showinfo("Success", "Employee list saved and updated successfully!")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to upload employee list: {e}")

def preview_dat_file(event):
    selected_index = employee_list_entry.curselection()
    if not selected_index:
        return

    file_path = employee_list_entry.get(selected_index[0])

    try:
        df = pd.read_csv(file_path, delimiter="\t", dtype=str)  # Read data file

        preview_window = tk.Toplevel(root)
        preview_window.title(f"Preview: {os.path.basename(file_path)}")
        preview_window.geometry("800x500")

        frame = tk.Frame(preview_window)
        frame.pack(expand=True, fill="both", padx=10, pady=10)

        tree = ttk.Treeview(frame, show="headings")

        # Set up columns
        tree["columns"] = list(df.columns)
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center", width=max(df[col].astype(str).str.len().max() * 8, 100))  # Auto adjust width

        # Insert rows
        for _, row in df.iterrows():
            tree.insert("", "end", values=list(row))

        # Add scrollbar
        scrollbar_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        scrollbar_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)

        scrollbar_x.pack(side="bottom", fill="x")
        scrollbar_y.pack(side="right", fill="y")
        tree.pack(expand=True, fill="both")

        # Add Convert button for DAT files
        def proceed_to_conversion():
            preview_window.destroy()
            convert_batch_to_excel([file_path])

        proceed_button = tk.Button(preview_window, text="Convert", command=proceed_to_conversion, 
                                  font=("Segoe UI", 12, "bold"), fg="white", bg="#4CAF50", 
                                  relief="flat", padx=10, pady=5)
        proceed_button.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")

def preview_txt_file(event):
    selected_index = employee_list_entry.curselection()
    if not selected_index:
        return

    file_path = employee_list_entry.get(selected_index[0])

    try:
        with open(file_path, 'r') as file:
            content = file.readlines()

        preview_window = tk.Toplevel(root)
        preview_window.title(f"Preview: {os.path.basename(file_path)}")
        preview_window.geometry("600x400")

        frame = tk.Frame(preview_window)
        frame.pack(expand=True, fill="both", padx=10, pady=10)

        # Create a Text widget to display the content
        text_widget = tk.Text(frame, wrap="none", font=("Courier New", 12))
        
        # Add scrollbars
        scrollbar_y = tk.Scrollbar(frame, orient="vertical", command=text_widget.yview)
        scrollbar_x = tk.Scrollbar(frame, orient="horizontal", command=text_widget.xview)
        text_widget.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Insert content
        for line in content:
            text_widget.insert(tk.END, line)
        
        # Make it read-only
        text_widget.config(state="disabled")
        
        # Arrange widgets
        scrollbar_y.pack(side="right", fill="y")
        scrollbar_x.pack(side="bottom", fill="x")
        text_widget.pack(side="left", fill="both", expand=True)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")

def save_to_database(filename, output_path):
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO conversions (filename, converted_at, output_path) VALUES (?, ?, ?)", 
                   (filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), output_path))
    conn.commit()
    conn.close()
    
    # Add to Excel listbox
    employee_list_entry.insert(tk.END, output_path)

def load_excel_history():
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("SELECT output_path FROM conversions ORDER BY id DESC")
    records = cursor.fetchall()
    conn.close()
    
    # Add existing Excel files to the listbox
    for record in records:
        if os.path.exists(record[0]):  # Only add if the file still exists
            employee_list_entry.insert(tk.END, record[0])

def open_excel_file(event):
    selected_index = employee_list_entry.curselection()
    if not selected_index:
        return

    file_path = employee_list_entry.get(selected_index[0])
    
    if os.path.exists(file_path):
        try:
            subprocess.run(["start", "", file_path], shell=True)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {e}")
    else:
        messagebox.showerror("Error", "File not found. It may have been moved or deleted.")
        # Remove entry from listbox if file doesn't exist
        employee_list_entry.delete(selected_index)

def browse_and_preview_employee_list():
    file_path = filedialog.askopenfilename(
        title="Select Employee List TXT File",
        filetypes=[("Text Files", "*.txt")]
    )

    if file_path:
        # Clear previous entries
        employee_list_entry.delete(0, tk.END)
        employee_list_entry.insert(0, file_path)

        preview_employee_list.config(state=tk.NORMAL)
        preview_employee_list.delete(1.0, tk.END)
        
        try:
            with open(file_path, 'r') as file:
                content = file.read()
                preview_employee_list.insert(tk.END, content)
        except Exception as e:
            preview_employee_list.insert(tk.END, f"Error reading file: {e}")
        
        preview_employee_list.config(state=tk.DISABLED)

def browse_and_preview_dat_files():
    files = filedialog.askopenfilenames(
        title="Select DAT Files",
        filetypes=[("DAT Files", "*.dat")]
    )

    if files:
        dat_file_entry.delete(0, tk.END)
        dat_file_entry.insert(0, ", ".join(files))

        preview_dat_files.config(state=tk.NORMAL)
        preview_dat_files.delete(1.0, tk.END)
        
        for file_path in files:
            try:
                with open(file_path, 'r') as file:
                    content = file.read()
                    preview_dat_files.insert(tk.END, f"File: {os.path.basename(file_path)}\n")
                    preview_dat_files.insert(tk.END, content + "\n\n")
            except Exception as e:
                preview_dat_files.insert(tk.END, f"Error reading {file_path}: {e}\n\n")
        
        preview_dat_files.config(state=tk.DISABLED)
def convert_files():
    employee_file = employee_list_entry.get()
    dat_files = dat_file_entry.get().split(", ")

    if not employee_file or not dat_files:
        messagebox.showwarning("Warning", "Please upload both Employee List and DAT Files")
        return

    try:
        # Update employee list first
        upload_employee_list_from_path(employee_file)

        # Convert DAT files
        convert_batch_to_excel(dat_files)

    except Exception as e:
        messagebox.showerror("Error", str(e))

def upload_employee_list_from_path(file_path):
    global employee_list
    try:
        conn = sqlite3.connect("conversion_history.db")
        cursor = conn.cursor()
        cursor.execute("DELETE FROM employees")  # Clear existing list

        with open(file_path, 'r') as file:
            employee_list.clear()
            for line in file:
                parts = line.strip().split(' ', 1)
                if len(parts) == 2:
                    emp_id = int(parts[0])
                    emp_name = parts[1].strip().upper()
                    employee_list[emp_id] = emp_name
                    cursor.execute("INSERT INTO employees (id, name) VALUES (?, ?)", (emp_id, emp_name))

        conn.commit()
        conn.close()
        messagebox.showinfo("Success", "Employee list updated successfully!")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to upload employee list: {e}")

class StyledTkinter:
    # Color Palette
    COLORS = {
        'bg_primary': '#f8f9fa',      # Light gray-white
        'bg_secondary': '#f1f3f5',    # Slightly darker white
        'bg_accent': '#e9ecef',       # Light gray accent

        'text_primary': '#1a365d',     # Dark navy blue
        'text_secondary': '#2c3e50',  # Slightly lighter navy blue
        'btn_primary': '#3182ce',     # Vibrant blue
        'btn_secondary': '#4a5568',   # Dark grayish blue
        'btn_success': '#48bb78',     # Green for success
        'btn_warning': '#ed8936',     # Orange for warnings
        'border_color': '#cbd5e0'     # Light border color
    }

    @classmethod
    def create_styled_button(cls, parent, text, command, style='primary', width=None):
        btn_styles = {
            'primary': {
                'bg': cls.COLORS['btn_primary'], 
                'fg': 'white', 
                'hover_bg': '#2c5282'
            },
            'secondary': {
                'bg': cls.COLORS['btn_secondary'], 
                'fg': 'white', 
                'hover_bg': '#718096'
            },
            'success': {
                'bg': cls.COLORS['btn_success'], 
                'fg': 'white', 
                'hover_bg': '#38a169'
            },
            'warning': {
                'bg': cls.COLORS['btn_warning'], 
                'fg': 'white', 
                'hover_bg': '#dd6b20'
            }
        }

        current_style = btn_styles.get(style, btn_styles['primary'])
        
        btn = tk.Button(
            parent, 
            text=text, 
            command=command,
            bg=current_style['bg'], 
            fg=current_style['fg'],
            font=("Segoe UI", 10, "bold"),
            relief=tk.FLAT,
            padx=10,
            pady=5
        )

        if width:
            btn.config(width=width)

        # Hover effects
        def on_enter(e):
            btn.config(bg=current_style['hover_bg'])

        def on_leave(e):
            btn.config(bg=current_style['bg'])

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

        return btn

    @classmethod
    def create_styled_entry(cls, parent, width=50):
        entry = tk.Entry(
            parent, 
            font=("Segoe UI", 10), 
            width=width,
            bg=cls.COLORS['bg_secondary'],
            fg=cls.COLORS['text_primary'],
            insertbackground=cls.COLORS['text_primary'],
            relief=tk.FLAT,
            highlightthickness=1,
            highlightcolor=cls.COLORS['border_color'],
            highlightbackground=cls.COLORS['border_color']
        )
        return entry

    @classmethod
    def create_styled_label(cls, parent, text, style='primary'):
        label_styles = {
            'primary': {
                'fg': cls.COLORS['text_primary'],
                'font': ("Segoe UI", 10, "bold")
            },
            'secondary': {
                'fg': cls.COLORS['text_secondary'],
                'font': ("Segoe UI", 10)
            }
        }
        
        current_style = label_styles.get(style, label_styles['primary'])
        
        label = tk.Label(
            parent, 
            text=text, 
            fg=current_style['fg'],
            font=current_style['font'],
            bg=cls.COLORS['bg_primary']
        )
        return label

def create_improved_gui():
    root = tk.Tk()
    root.title("Employee DTR Converter")
    root.geometry("900x800")
    root.configure(bg=StyledTkinter.COLORS['bg_primary'])

    # Main Container
    main_container = tk.Frame(root, bg=StyledTkinter.COLORS['bg_primary'], padx=30, pady=30)
    main_container.pack(fill=tk.BOTH, expand=True)

    # Title
    title_label = StyledTkinter.create_styled_label(
        main_container, 
        "Employee DTR Converter", 
        style='primary'
    )
    title_label.config(font=("Segoe UI", 18, "bold"))
    title_label.pack(pady=(0, 20))

    # Employee List Section
    employee_section = tk.LabelFrame(
        main_container, 
        text="Employee List", 
        font=("Segoe UI", 12, "bold"), 
        bg=StyledTkinter.COLORS['bg_primary'],
        fg=StyledTkinter.COLORS['text_primary'],
        labelanchor='n',
        borderwidth=2,
        relief=tk.GROOVE
    )
    employee_section.pack(fill=tk.X, pady=10)

    employee_row = tk.Frame(employee_section, bg=StyledTkinter.COLORS['bg_primary'])
    employee_row.pack(padx=10, pady=10, fill=tk.X)

    # Employee List Label
    employee_list_label = StyledTkinter.create_styled_label(
        employee_row, 
        "Employee List File:", 
        style='secondary'
    )
    employee_list_label.pack(side=tk.LEFT, padx=(0, 10))

    # Employee List Entry
    employee_list_entry = StyledTkinter.create_styled_entry(employee_row, width=50)
    employee_list_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

    # Employee List Browse Button
    employee_upload_btn = StyledTkinter.create_styled_button(
        employee_row, 
        "Browse", 
        lambda: browse_and_preview_employee_list(),
        style='primary'
    )
    employee_upload_btn.pack(side=tk.LEFT)

    # Employee List Preview Frame
    employee_preview_frame = tk.Frame(employee_section, bg=StyledTkinter.COLORS['bg_primary'])
    employee_preview_frame.pack(padx=10, pady=10, fill=tk.X)

    # Employee List Preview
    preview_employee_list = tk.Text(
        employee_preview_frame, 
        height=5, 
        font=("Courier", 10), 
        state=tk.DISABLED, 
        wrap=tk.NONE,
        bg=StyledTkinter.COLORS['bg_secondary'],
        fg=StyledTkinter.COLORS['text_primary']
    )
    preview_employee_list.pack(side=tk.LEFT, expand=True, fill=tk.X)

    # Employee Preview Scrollbars
    employee_preview_scrollbar_y = tk.Scrollbar(
        employee_preview_frame, 
        orient=tk.VERTICAL, 
        command=preview_employee_list.yview,
        bg=StyledTkinter.COLORS['bg_accent']
    )
    employee_preview_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    
    employee_preview_scrollbar_x = tk.Scrollbar(
        employee_section, 
        orient=tk.HORIZONTAL, 
        command=preview_employee_list.xview,
        bg=StyledTkinter.COLORS['bg_accent']
    )
    employee_preview_scrollbar_x.pack(fill=tk.X, padx=10)
    
    preview_employee_list.configure(
        yscrollcommand=employee_preview_scrollbar_y.set,
        xscrollcommand=employee_preview_scrollbar_x.set
    )

    # DAT Files Section
    dat_section = tk.LabelFrame(
        main_container, 
        text="DTR Files", 
        font=("Segoe UI", 12, "bold"), 
        bg=StyledTkinter.COLORS['bg_primary'],
        fg=StyledTkinter.COLORS['text_primary'],
        labelanchor='n',
        borderwidth=2,
        relief=tk.GROOVE
    )
    dat_section.pack(fill=tk.X, pady=10)

    dat_row = tk.Frame(dat_section, bg=StyledTkinter.COLORS['bg_primary'])
    dat_row.pack(padx=10, pady=10, fill=tk.X)

    # DAT Files Label
    dat_file_label = StyledTkinter.create_styled_label(
        dat_row, 
        "DTR Files:", 
        style='secondary'
    )
    dat_file_label.pack(side=tk.LEFT, padx=(0, 10))

    # DAT Files Entry
    dat_file_entry = StyledTkinter.create_styled_entry(dat_row, width=50)
    dat_file_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

    # DAT Files Browse Button
    dat_upload_btn = StyledTkinter.create_styled_button(
        dat_row, 
        "Browse", 
        lambda: browse_and_preview_dat_files(),
        style='primary'
    )
    dat_upload_btn.pack(side=tk.LEFT)

    # DAT Files Preview Frame
    dat_preview_frame = tk.Frame(dat_section, bg=StyledTkinter.COLORS['bg_primary'])
    dat_preview_frame.pack(padx=10, pady=10, fill=tk.X)

    # DAT Files Preview
    preview_dat_files = tk.Text(
        dat_preview_frame, 
        height=10, 
        font=("Courier", 10), 
        state=tk.DISABLED, 
        wrap=tk.NONE,
        bg=StyledTkinter.COLORS['bg_secondary'],
        fg=StyledTkinter.COLORS['text_primary']
    )
    preview_dat_files.pack(side=tk.LEFT, expand=True, fill=tk.X)

    # DAT Preview Scrollbars
    dat_preview_scrollbar_y = tk.Scrollbar(
        dat_preview_frame, 
        orient=tk.VERTICAL, 
        command=preview_dat_files.yview,
        bg=StyledTkinter.COLORS['bg_accent']
    )
    dat_preview_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
    
    dat_preview_scrollbar_x = tk.Scrollbar(
        dat_section, 
        orient=tk.HORIZONTAL, 
        command=preview_dat_files.xview,
        bg=StyledTkinter.COLORS['bg_accent']
    )
    dat_preview_scrollbar_x.pack(fill=tk.X, padx=10)
    
    preview_dat_files.configure(
        yscrollcommand=dat_preview_scrollbar_y.set,
        xscrollcommand=dat_preview_scrollbar_x.set
    )

    # Conversion and History Buttons
    button_frame = tk.Frame(main_container, bg=StyledTkinter.COLORS['bg_primary'])
    button_frame.pack(pady=10)

    history_btn = StyledTkinter.create_styled_button(
        button_frame, 
        "View Conversion History", 
        show_history,
        style='secondary'
    )
    history_btn.pack(side=tk.LEFT, padx=10)

    convert_btn = StyledTkinter.create_styled_button(
        button_frame, 
        "Convert DTR to Excel", 
        convert_files,
        style='success'
    )
    convert_btn.pack(side=tk.LEFT, padx=10)

    return root, employee_list_entry, preview_employee_list, dat_file_entry, preview_dat_files
    

# Replace the original create_gui() function with this new one
# Modify the main block
if __name__ == "__main__":
    create_database()
    load_employee_list()  # Load employee list from DB if any
    
    # Create the root window and key widgets using the new method
    root, employee_list_entry, preview_employee_list, dat_file_entry, preview_dat_files = create_improved_gui()
    
    # Update global references if needed
    globals()['employee_list_entry'] = employee_list_entry
    globals()['preview_employee_list'] = preview_employee_list
    globals()['dat_file_entry'] = dat_file_entry
    globals()['preview_dat_files'] = preview_dat_files
    
    root.mainloop()