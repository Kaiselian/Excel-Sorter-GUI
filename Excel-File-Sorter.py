import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def upload_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    if file_path:
        try:
            if file_path.endswith(".csv"):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)
            messagebox.showinfo("Success", "File uploaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")


def search_and_generate():
    if df is None:
        messagebox.showerror("Error", "Please upload a file first.")
        return

    query = search_var.get().strip()
    if not query:
        messagebox.showerror("Error", "Please enter an ID or name to search.")
        return

    filtered_data = df[df.apply(lambda row: row.astype(str).str.contains(query, case=False, na=False).any(), axis=1)]

    if filtered_data.empty:
        messagebox.showinfo("No Results", "No matching records found.")
        return

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        try:
            filtered_data.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "Filtered data saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")


# GUI Setup
root = tk.Tk()
root.title("Data Search & Export Tool")
root.geometry("1280x720")

df = None  # To store the uploaded file

# Upload Button
upload_btn = tk.Button(root, text="Upload File", command=upload_file)
upload_btn.pack(pady=25)

# Search Bar
search_var = tk.StringVar()
search_entry = tk.Entry(root, textvariable=search_var, width=40)
search_entry.pack(pady=15)

# Generate Button
generate_btn = tk.Button(root, text="Generate Excel File", command=search_and_generate)
generate_btn.pack(pady=25)

root.mainloop()
