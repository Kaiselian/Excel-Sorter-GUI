import ttkbootstrap as tb
from ttkbootstrap.constants import *
import pandas as pd
from tkinter import filedialog, messagebox

# Function to Upload File
def upload_file():
    global df
    file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xls"), ("CSV files", "*.csv")])
    if file_path:
        try:
            if file_path.endswith(".csv"):
                df = pd.read_csv(file_path)
            else:
                df = pd.read_excel(file_path)

            display_data(df)
            messagebox.showinfo("Success", "File uploaded successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

# Function to Display Data in Treeview
def display_data(data):
    for row in tree.get_children():
        tree.delete(row)

    tree["columns"] = list(data.columns)
    tree["show"] = "headings"

    for col in data.columns:
        tree.heading(col, text=col, anchor="center")
        tree.column(col, width=120, anchor="center")

    for _, row in data.iterrows():
        tree.insert("", "end", values=list(row))

# Function to Search and Generate Excel File
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

    display_data(filtered_data)

    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if save_path:
        try:
            filtered_data.to_excel(save_path, index=False)
            messagebox.showinfo("Success", "Filtered data saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {e}")

# GUI Setup
root = tb.Window(themename="darkly")  # Modern dark theme
root.title("Modern Data Search & Export Tool")
root.geometry("1280x720")

df = None  # Global DataFrame

# Frame for Top Controls
top_frame = tb.Frame(root)
top_frame.pack(pady=10, fill=X, padx=20)

upload_btn = tb.Button(top_frame, text="Upload File", bootstyle=PRIMARY, command=upload_file)
upload_btn.pack(side=LEFT, padx=10)

search_var = tb.StringVar()
search_entry = tb.Entry(top_frame, textvariable=search_var, width=40)
search_entry.pack(side=LEFT, padx=10)

search_btn = tb.Button(top_frame, text="Search", bootstyle=SUCCESS, command=search_and_generate)
search_btn.pack(side=LEFT, padx=10)

# Treeview for Data Display
tree_frame = tb.Frame(root)
tree_frame.pack(pady=10, fill=BOTH, expand=True, padx=20)

tree = tb.Treeview(tree_frame, bootstyle="info")
tree.pack(fill=BOTH, expand=True)

# Generate Button
generate_btn = tb.Button(root, text="Generate Excel File", bootstyle=WARNING, command=search_and_generate)
generate_btn.pack(pady=10)

root.mainloop()
