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

def create_improved_gui():
    global root, employee_list_entry, dat_file_entry, preview_employee_list, preview_dat_files

    root = tk.Tk()
    root.title("Employee DTR Converter")
    root.geometry("800x700")
    root.configure(bg="#f0f0f0")

    # Main Container
    main_container = tk.Frame(root, bg="#f0f0f0", padx=20, pady=20)
    main_container.pack(fill=tk.BOTH, expand=True)

    # Title
    title_label = tk.Label(main_container, text="Employee DTR Converter", 
                            font=("Segoe UI", 15, "bold"), bg="#f0f0f0", fg="#333")
    title_label.pack(pady=(0, 10))

    # Employee List Section
    employee_section = tk.LabelFrame(main_container, text="Employee List", 
                                     font=("Segoe UI", 12, "bold"), bg="#f0f0f0")
    employee_section.pack(fill=tk.X, pady=10)

    employee_row = tk.Frame(employee_section, bg="#f0f0f0")
    employee_row.pack(padx=10, pady=10, fill=tk.X)

    employee_list_label = tk.Label(employee_row, text="Employee List File:", 
                                   font=("Segoe UI", 10), bg="#f0f0f0")
    employee_list_label.pack(side=tk.LEFT, padx=(0, 10))

    employee_list_entry = tk.Entry(employee_row, font=("Segoe UI", 10), width=50)
    employee_list_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

    employee_upload_btn = tk.Button(employee_row, text="Browse", 
                                    command=lambda: browse_and_preview_employee_list(),
                                    font=("Segoe UI", 10), bg="#4CAF50", fg="white")
    employee_upload_btn.pack(side=tk.LEFT)

    # Employee List Preview
    preview_employee_list = tk.Text(employee_section, height=5, font=("Courier", 10), 
                                    state=tk.DISABLED, wrap=tk.NONE)
    preview_employee_list.pack(padx=10, pady=10, fill=tk.X)

    # DAT Files Section
    dat_section = tk.LabelFrame(main_container, text="DTR Files", 
                                font=("Segoe UI", 12, "bold"), bg="#f0f0f0")
    dat_section.pack(fill=tk.X, pady=10)

    dat_row = tk.Frame(dat_section, bg="#f0f0f0")
    dat_row.pack(padx=10, pady=10, fill=tk.X)

    dat_file_label = tk.Label(dat_row, text="DTR Files:", 
                              font=("Segoe UI", 10), bg="#f0f0f0")
    dat_file_label.pack(side=tk.LEFT, padx=(0, 10))

    dat_file_entry = tk.Entry(dat_row, font=("Segoe UI", 10), width=50)
    dat_file_entry.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 10))

    dat_upload_btn = tk.Button(dat_row, text="Browse", 
                               command=lambda: browse_and_preview_dat_files(),
                               font=("Segoe UI", 10), bg="#2196F3", fg="white")
    dat_upload_btn.pack(side=tk.LEFT)

    # DAT Files Preview
    preview_dat_files = tk.Text(dat_section, height=10, font=("Courier", 10), 
                                state=tk.DISABLED, wrap=tk.NONE)
    preview_dat_files.pack(padx=10, pady=10, fill=tk.X)

    # History Button
    history_btn = tk.Button(main_container, text="View Conversion History", 
                            command=show_history, 
                            font=("Segoe UI", 12), 
                            bg="#9C27B0", fg="white")
    history_btn.pack(pady=2)

    # Conversion Button
    convert_btn = tk.Button(main_container, text="Convert DTR to Excel", 
                            command=convert_files, 
                            font=("Segoe UI", 12, "bold"), 
                            bg="#FF9800", fg="white")
    convert_btn.pack(pady=2, padx=20)

    # Scroll bars for previews
    employee_preview_scrollbar_x = tk.Scrollbar(employee_section, orient=tk.HORIZONTAL, 
                                                command=preview_employee_list.xview)
    employee_preview_scrollbar_x.pack(fill=tk.X, padx=3)
    preview_employee_list.configure(xscrollcommand=employee_preview_scrollbar_x.set)

    dat_preview_scrollbar_x = tk.Scrollbar(dat_section, orient=tk.HORIZONTAL, 
                                           command=preview_dat_files.xview)
    dat_preview_scrollbar_x.pack(fill=tk.X, padx=10)
    preview_dat_files.configure(xscrollcommand=dat_preview_scrollbar_x.set)

    root.mainloop()

# Replace the original create_gui() function with this new one
# Modify the main block
if __name__ == "__main__":
    create_database()
    load_employee_list()  # Load employee list from DB if any
    create_improved_gui()