import os
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

# Name mapping dictionary
name_mapping = {
    0: "Norman", 8: "Kian", 1: "Alice", 2: "Bob", 3: "Charlie",
    4: "David", 5: "Emma", 6: "Fiona", 7: "George", 9: "Henry"
}

# Global variables
history_window_open = False
history_window = None  # Ensure global reference to history window

# Create or connect to the database
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

# Save conversion details to database
def save_to_database(filename, output_path):
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("INSERT INTO conversions (filename, converted_at, output_path) VALUES (?, ?, ?)", 
                   (filename, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), output_path))
    conn.commit()
    conn.close()

# Function to filter in/out entries (ensures only first and last entry per day per employee)
def filter_in_out_entries(df):
    if df.shape[1] < 2:
        return df  # Return original if insufficient columns

    first_col = df.columns[0]  # Employee ID column
    time_col = df.columns[1]  # Timestamp column

    # Ensure the time column is in datetime format
    df[time_col] = pd.to_datetime(df[time_col])

    # Extract the date for grouping
    df['Date'] = df[time_col].dt.date

    # Find the first and last occurrence per employee per day
    first_occurrence = df.groupby([first_col, 'Date'])[time_col].idxmin()
    last_occurrence = df.groupby([first_col, 'Date'])[time_col].idxmax()

    # Combine both occurrences and restore original order based on date
    unique_indices = sorted(set(first_occurrence) | set(last_occurrence), key=lambda x: df.loc[x, time_col])
    filtered_df = df.loc[unique_indices].reset_index(drop=True)

    return filtered_df.drop(columns=['Date'])  # Remove temporary Date column


# Convert multiple .dat files to Excel with custom save location and filename
def convert_batch_to_excel(files):
    for dat_file in files:
        try:
            df = pd.read_csv(dat_file, delimiter="\t")

            if df.shape[1] > 0:
                first_column_name = df.columns[0]
                df[first_column_name] = df[first_column_name].map(name_mapping).fillna(df[first_column_name])

                # Apply filter to keep only first and last entry per employee per day
                df = filter_in_out_entries(df)

            # Ask user for save location and custom filename
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=os.path.basename(dat_file).replace(".dat", ".xlsx"),
                title="Save Converted Excel File"
            )

            if save_path:  # Proceed only if the user selected a path
                df.to_excel(save_path, index=False, engine='openpyxl')
                save_to_database(os.path.basename(dat_file), save_path)
            else:
                messagebox.showwarning("Warning", "File saving canceled.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to convert {os.path.basename(dat_file)}: {e}")
    
    messagebox.showinfo("Success", "Batch conversion completed!")

# Retrieve and display conversion history with search functionality
def show_history():
    global history_window, history_window_open

    if not history_window_open:
        history_window_open = True  # Mark history window as open

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

        history_window.protocol("WM_DELETE_WINDOW", close_history_window)  # Close history window

    else:
        messagebox.showinfo("Info", "History window is already open.")

# Close history window and reset global flag
def close_history_window():
    global history_window, history_window_open
    if history_window:
        history_window.destroy()
        history_window = None  # Reset history window reference
    history_window_open = False

# Browse multiple .dat files
def browse_files():
    files = filedialog.askopenfilenames(filetypes=[("Data Files", "*.dat")])
    if files:
        convert_batch_to_excel(files)

# Create GUI
def create_gui():
    global root, label, frame, button, history_button
    root = tk.Tk()
    root.title("DAT to Excel Converter")
    root.geometry("400x350")  # You can adjust the initial size
    root.configure(bg="#f5f5f5")
    root.resizable(True, True)  # Enable resizing (minimize and maximize)

    try:
        root.iconbitmap("edplogo.ico")  
    except:
        print("Icon not found, using default")

    label = tk.Label(root, text="Convert DAT to Excel", font=("Segoe UI", 18, "bold"), bg="#f5f5f5", fg="#333")
    label.pack(pady=20)

    frame = tk.Frame(root, bg="#f5f5f5")
    frame.pack(pady=10)

    button = tk.Button(frame, text="Select .dat Files", command=browse_files, font=("Segoe UI", 12, "bold"), fg="white", bg="#4CAF50", relief="flat", bd=0, padx=20, pady=10, highlightthickness=0)
    button.grid(row=0, column=0)

    history_button = tk.Button(frame, text="View History", command=show_history, font=("Segoe UI", 12, "bold"), fg="white", bg="#2196F3", relief="flat", bd=0, padx=20, pady=10, highlightthickness=0)
    history_button.grid(row=1, column=0, pady=10)

    root.mainloop()

if __name__ == "__main__":
    create_database()
    create_gui()
