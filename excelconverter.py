import os
import sqlite3
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from datetime import datetime

# Name mapping dictionary
name_mapping = {
    0: "Norman",
    8: "Kian",
    1: "Alice",
    2: "Bob",
    3: "Charlie",
    4: "David",
    5: "Emma",
    6: "Fiona",
    7: "George",
    9: "Henry"
}

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

# Retrieve and display conversion history
def show_history():
    conn = sqlite3.connect("conversion_history.db")
    cursor = conn.cursor()
    cursor.execute("SELECT filename, converted_at, output_path FROM conversions ORDER BY id DESC")
    records = cursor.fetchall()
    conn.close()

    history_window = tk.Toplevel(root)
    history_window.title("Conversion History")
    history_window.geometry("500x300")
    history_window.configure(bg="#f5f5f5")

    tree = ttk.Treeview(history_window, columns=("Filename", "Date", "Output Path"), show="headings")
    tree.heading("Filename", text="Filename")
    tree.heading("Date", text="Date Converted")
    tree.heading("Output Path", text="Output Path")
    tree.column("Filename", width=150)
    tree.column("Date", width=120)
    tree.column("Output Path", width=200)

    for record in records:
        tree.insert("", "end", values=record)

    tree.pack(expand=True, fill="both")

# Function to convert .dat file to Excel and replace numbers with names
def convert_to_excel(dat_file):
    try:
        df = pd.read_csv(dat_file, delimiter="\t")  # Adjust delimiter if needed
        
        # Check if the first column exists
        if df.shape[1] > 0:
            first_column_name = df.columns[0]  # Get the first column name
            
            # Replace numbers in the first column using the name mapping dictionary
            df[first_column_name] = df[first_column_name].map(name_mapping).fillna(df[first_column_name])

        excel_file = os.path.splitext(dat_file)[0] + '.xlsx'
        df.to_excel(excel_file, index=False, engine='openpyxl')
        save_to_database(os.path.basename(dat_file), excel_file)  # Save to database
        messagebox.showinfo("Success", f"File converted successfully!\nSaved as: {excel_file}")
        download_file(excel_file)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to convert file: {e}")

# Function to handle the file download (move to a user-selected location)
def download_file(excel_file):
    download_path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=os.path.basename(excel_file), filetypes=[("Excel Files", "*.xlsx")])
    if download_path:
        try:
            os.rename(excel_file, download_path)
            messagebox.showinfo("Download", f"File successfully downloaded to: {download_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

# Function to browse and select .dat file
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Data Files", "*.dat")])
    if file_path:
        convert_to_excel(file_path)

# Function to change button appearance on hover
def on_enter(event, button):
    button.config(bg="#45a049")

def on_leave(event, button):
    button.config(bg="#4CAF50")

# Create the GUI for the application
def create_gui():
    global root
    root = tk.Tk()
    root.title("DAT to Excel Converter")
    root.geometry("400x350")
    root.configure(bg="#f5f5f5")
    root.resizable(False, False)

    try:
        root.iconbitmap("app_icon.ico")  
    except:
        print("Icon not found, using default")

    label = tk.Label(root, text="Convert DAT to Excel", font=("Segoe UI", 18, "bold"), bg="#f5f5f5", fg="#333")
    label.pack(pady=20)

    frame = tk.Frame(root, bg="#f5f5f5")
    frame.pack(pady=10)

    button = tk.Button(
        frame, text="Select .dat File", command=browse_file,
        font=("Segoe UI", 12, "bold"), fg="white", bg="#4CAF50",
        relief="flat", bd=0, padx=20, pady=10, highlightthickness=0,
        activebackground="#45a049"
    )
    button.grid(row=0, column=0)

    button.bind("<Enter>", lambda e, button=button: on_enter(e, button))
    button.bind("<Leave>", lambda e, button=button: on_leave(e, button))

    history_button = tk.Button(
        frame, text="View History", command=show_history,
        font=("Segoe UI", 12, "bold"), fg="white", bg="#2196F3",
        relief="flat", bd=0, padx=20, pady=10, highlightthickness=0,
        activebackground="#1976D2"
    )
    history_button.grid(row=1, column=0, pady=10)

    history_button.bind("<Enter>", lambda e, button=history_button: button.config(bg="#1976D2"))
    history_button.bind("<Leave>", lambda e, button=history_button: button.config(bg="#2196F3"))

    root.mainloop()

if __name__ == "__main__":
    create_database()
    create_gui()
