import os
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime
import subprocess

# Name mapping dictionary
name_mapping = {
    0: "Norman", 8: "Kian", 1: "Alice", 2: "Bob", 3: "Charlie",
    4: "David", 5: "Emma", 6: "Fiona", 7: "George", 9: "Henry"
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
    
    first_col = df.columns[0]  # Employee ID column
    time_col = df.columns[1]  # Timestamp column

    df[time_col] = pd.to_datetime(df[time_col])
    df['Date'] = df[time_col].dt.date

    first_occurrence = df.groupby([first_col, 'Date'])[time_col].idxmin()
    last_occurrence = df.groupby([first_col, 'Date'])[time_col].idxmax()

    unique_indices = sorted(set(first_occurrence) | set(last_occurrence), key=lambda x: df.loc[x, time_col])
    filtered_df = df.loc[unique_indices].reset_index(drop=True)
    
    return filtered_df.drop(columns=['Date'])  # Remove temporary column

def convert_batch_to_excel(files):
    for dat_file in files:
        try:
            df = pd.read_csv(dat_file, delimiter="\t")
            
            if df.shape[1] > 0:
                first_column_name = df.columns[0]
                df[first_column_name] = df[first_column_name].map(name_mapping).fillna(df[first_column_name])
                df = filter_in_out_entries(df)
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=os.path.basename(dat_file).replace(".dat", ".xlsx"),
                title="Save Converted Excel File"
            )
            
            if save_path:
                df.to_excel(save_path, index=False, engine='openpyxl')
                save_to_database(os.path.basename(dat_file), save_path)
                
                # Ask if user wants to open the file
                open_file = messagebox.askyesno("Conversion Complete", "File converted successfully!\nDo you want to open it now?")
                if open_file:
                    subprocess.run(["start", "", save_path], shell=True)

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert {os.path.basename(dat_file)}: {e}")

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

    try:
        df = pd.read_csv(file_path, delimiter="\t", dtype=str)  # Ensure all columns are read properly

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
    global root
    root = tk.Tk()
    root.title("DAT to Excel Converter")
    root.geometry("400x350")
    root.configure(bg="#f5f5f5")
    root.resizable(True, True)
    
    label = tk.Label(root, text="Convert DAT to Excel", font=("Segoe UI", 18, "bold"), bg="#f5f5f5", fg="#333")
    label.pack(pady=20)

    frame = tk.Frame(root, bg="#f5f5f5")
    frame.pack(pady=10)

    button = tk.Button(frame, text="Select .dat Files", command=browse_files, font=("Segoe UI", 12, "bold"), fg="white", bg="#4CAF50", relief="flat", padx=20, pady=10)
    button.grid(row=0, column=0)

    history_button = tk.Button(frame, text="View History", command=show_history, font=("Segoe UI", 12, "bold"), fg="white", bg="#2196F3", relief="flat", padx=20, pady=10)
    history_button.grid(row=1, column=0, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_database()
    create_gui()
