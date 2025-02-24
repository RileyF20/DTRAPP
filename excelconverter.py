import os
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter
from openpyxl import load_workbook
import subprocess
import openpyxl

# Name mapping dictionary
name_mapping = {
    1: "PALILEO, FORTUNATO L.",
    2: "QUILANTANG, ALLAN C.",
    3: "BONDOC, FREDERICK B.",
    4: "TAMAYO, ROBERT G.",
    5: "ALFONSO, ROSAMIE Y.",
    6: "DE LEON, ESPERANZA M.",
    7: "MANIQUIS, MA. ANGELINA B.",
    8: "JACINTO, FERNANDO JR. N.",
    9: "ILETO, ELIZABETH C.",
    11: "PEÑA, JOHN RONWALDO C.",
    12: "DEPANO, ELSON JR. T.",
    13: "OMBAO, RODEL A.",
    14: "VELINA, TIBOY JR. M.",
    15: "SUBIDO, MARIETTA E.",
    16: "CABANDING, TERESITA C.",
    17: "GURION, CHRISTOPHER T.",
    18: "PASCUA, NERISSA D.",
    21: "PAZ, CHARMAINE T.",
    22: "BUCAYU, NORMA F.",
    23: "GUNGON, JAMES CHRISTIAN N.",
    24: "ESTEVES, ALLAN M.",
    25: "POMAREJOS, KATHLEEN C.",
    27: "GOPEZ, RICHARD M.",
    29: "REYES, JOSE GLENN G.",
    30: "CRUZ, ARIEL C.",
    32: "ALMERO, BEVERLY O.",
    36: "ALZAGA, EMY F.",
    37: "RAYOS, FERNANDO G.",
    38: "BARING, ROBERT I.",
    39: "RIESGO, RECHELLE N.",
    40: "BERNALDO, MA. THERESA C.",
    60: "DOMINGO, KIM EDWARD B."
}

# Global variables
history_window_open = False
history_window = None  # Ensure global reference to history window

def create_database():
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS conversions (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            filename TEXT NOT NULL,
            converted_at TEXT NOT NULL,
            output_path TEXT NOT NULL
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

    first_col = df.columns[0]  # Employee ID or Name column
    time_col = df.columns[1]  # Timestamp column

    # Convert timestamp to datetime
    df[time_col] = pd.to_datetime(df[time_col])

    # Extract date and time
    df['Date'] = df[time_col].dt.strftime('%Y-%m-%d')  # Extract date
    df['Time'] = df[time_col].dt.strftime('%H:%M:%S')  # Extract time only

    # Get first time in and last time out for each employee per day
    grouped = df.groupby([first_col, 'Date'])['Time'].agg(['first', 'last']).reset_index()
    grouped.columns = ['Name', 'Date', 'Time In', 'Time Out']

    # Add employee number based on the name mapping
    reverse_mapping = {v: k for k, v in name_mapping.items()}  # Reverse mapping
    grouped['Employee No.'] = grouped['Name'].map(reverse_mapping)

    # If only one log exists, mark 'Time Out' as "No Out"
    grouped['Time Out'] = grouped.apply(
        lambda row: row['Time Out'] if row['Time In'] != row['Time Out'] else 'No Out', axis=1
    )

    # Pivot the data so each day has Time In and Time Out side by side
    result = grouped.pivot(index=['Employee No.', 'Name'], columns='Date', values=['Time In', 'Time Out'])
    result = result.swaplevel(axis=1).sort_index(axis=1, level=0)

    # Flatten MultiIndex columns for better formatting
    result.columns = [f"{date} {status}" for date, status in result.columns]
    
    # Reset index and sort based on the employee number
    result = result.reset_index()
    result = result.sort_values(by="Employee No.").reset_index(drop=True)

    return result


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
            df["Name"] = df["Name"].map(name_mapping).fillna(df["Name"].astype(str))

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

                open_file = messagebox.askyesno("Conversion Complete", "File converted successfully!\nDo you want to open it now?")
                if open_file:
                    subprocess.run(["start", "", save_path], shell=True)

                # Remove the .dat file from the listbox after successful conversion
                listbox_files.delete(listbox_files.get(0, tk.END).index(dat_file))  # Remove the file from the listbox

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert {os.path.basename(dat_file)}: {e}")


def generate_employee_dtr(writer, df, employee_name):
    """Generates a formatted Daily Time Record (DTR) sheet for an employee."""
    employee_df = df[df["Name"] == employee_name].copy()
    employee_df["Date"] = employee_df["Timestamp"].dt.date
    employee_df["Time"] = employee_df["Timestamp"].dt.strftime('%H:%M:%S')

    # Group by Date to get first and last times (AM Arrival, PM Departure)
    grouped = employee_df.groupby("Date")["Time"].agg(["first", "last"]).reset_index()
    grouped.columns = ["Date", "AM Arrival", "PM Departure"]

    # Mark "No Out" if there's only one timestamp
    grouped["PM Departure"] = grouped.apply(
        lambda row: row["PM Departure"] if row["AM Arrival"] != row["PM Departure"] else "No Out", axis=1
    )

    # Generate full month dates for missing records
    first_date = grouped["Date"].min()
    last_date = grouped["Date"].max()
    all_dates = pd.date_range(first_date, last_date, freq='D').date
    dtr_df = pd.DataFrame({"Date": all_dates})

    # Merge with recorded time logs
    dtr_df = dtr_df.merge(grouped, on="Date", how="left")

    # Extract day number and weekday names
    dtr_df["Day"] = dtr_df["Date"].apply(lambda x: x.day)
    dtr_df["Weekday"] = dtr_df["Date"].apply(lambda x: x.strftime('%A'))

    # Fill in "Sunday" and "Holiday" labels
    dtr_df["AM Arrival"].fillna(dtr_df["Weekday"].apply(lambda x: "Sunday" if x == "Sunday" else ""), inplace=True)
    dtr_df["PM Departure"].fillna(dtr_df["Weekday"].apply(lambda x: "Sunday" if x == "Sunday" else ""), inplace=True)

    # Reorder columns for DTR format and add Weekday, AM Arrival, PM Departure
    dtr_df = dtr_df[["Day", "Weekday", "AM Arrival", "PM Departure"]]

    # Write to Excel with formatting
    dtr_df.to_excel(writer, sheet_name=employee_name, index=False, startrow=4)

    # Apply DTR formatting in Excel
    wb = writer.book
    ws = wb[employee_name]

    # Add title and headers
    ws["A1"] = "DAILY TIME RECORD"
    ws["A2"] = f"Employee: {employee_name}"
    ws["A1"].font = Font(size=14, bold=True)
    ws["A2"].font = Font(size=12, italic=True)

    # Description section below the DTR
    month_year = first_date.strftime("%B %d-%d %Y")  # Example: "January 1-31 2025"
    ws["A3"] = f"For the month of: {month_year}"
    ws["A4"] = "Official hours for arrival and departure: "
    ws["A5"] = "Regular days: 8:00 AM - 5:00 PM"
    ws["A6"] = "Saturdays: "

    # Format description cells
    for row in range(3, 7):
        ws[f"A{row}"].font = Font(size=10, italic=True)
        ws[f"A{row}"].alignment = Alignment(horizontal="left")

    # Adjust column widths
    for col in ws.columns:
        max_length = 0
        col_letter = col[0].column_letter
        for cell in col:
            try:
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            except:
                pass
        ws.column_dimensions[col_letter].width = max_length + 2


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

def browse_files():
    file_path = filedialog.askopenfilename(filetypes=[("Data Files", "*.dat")])
    if not file_path:
        return

    listbox_files.insert(tk.END, file_path)  # Add filename to the listbox

def preview_selected_file(event):
    selected_index = listbox_files.curselection()
    if not selected_index:
        return

    file_path = listbox_files.get(selected_index[0])

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

        def proceed_to_conversion():
            preview_window.destroy()
            convert_batch_to_excel([file_path])

        proceed_button = tk.Button(preview_window, text="Convert", command=proceed_to_conversion, font=("Segoe UI", 12, "bold"), fg="white", bg="#4CAF50", relief="flat", padx=10, pady=5)
        proceed_button.pack(pady=10)

    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file: {e}")

def create_gui():
    global root, listbox_files
    root = tk.Tk()
    root.title("DAT to Excel Converter")
    root.geometry("500x400")
    root.configure(bg="#f5f5f5")
    root.resizable(True, True)

    label = tk.Label(root, text="Convert DAT to Excel", font=("Segoe UI", 18, "bold"), bg="#f5f5f5", fg="#333")
    label.pack(pady=20)

    frame = tk.Frame(root, bg="#f5f5f5")
    frame.pack(pady=10)

    button = tk.Button(frame, text="Select .dat Files", command=browse_files, font=("Segoe UI", 12, "bold"), 
                       fg="white", bg="#4CAF50", relief="flat", padx=20, pady=10, cursor="hand2")
    button.grid(row=0, column=0, padx=5)

    history_button = tk.Button(frame, text="View History", command=show_history, font=("Segoe UI", 12, "bold"), 
                               fg="white", bg="#2196F3", relief="flat", padx=20, pady=10, cursor="hand2")
    history_button.grid(row=0, column=1, padx=5)

    # Frame for Listbox and Scrollbar
    listbox_frame = tk.Frame(root, bg="#f5f5f5")
    listbox_frame.pack(padx=10, pady=10, fill="both", expand=True)

    # Scrollbar for the Listbox
    scrollbar = tk.Scrollbar(listbox_frame, orient="vertical")

    # Listbox with better UI styling
    listbox_files = tk.Listbox(listbox_frame, height=8, font=("Segoe UI", 12), 
                               selectbackground="#D3D3D3", relief="solid", bd=1, 
                               highlightthickness=1, highlightcolor="#4CAF50", yscrollcommand=scrollbar.set)
    listbox_files.pack(side="left", fill="both", expand=True)

    scrollbar.config(command=listbox_files.yview)
    scrollbar.pack(side="right", fill="y")

    # Bind double-click event to open preview
    listbox_files.bind("<Double-Button-1>", preview_selected_file)

    root.mainloop()

if __name__ == "__main__":
    create_gui()