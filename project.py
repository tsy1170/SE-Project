import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import os
import shutil
import barcode
from barcode.writer import ImageWriter
from datetime import datetime

tree = None
tree_frame = None

# def button_clicked(name):
#     # content_label.config(text=f"{name} button clicked!")

def view_all():
    global tree, tree_frame
    if tree:
        tree.destroy()

    if tree_frame:
        tree_frame.destroy()

    # Get script directory
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_files = [f for f in os.listdir(script_dir) if f.endswith((".xlsx", ".xls"))]

    if not excel_files:
        messagebox.showerror("Error", "No Excel files found.")
        return

    # Frame to hold the treeview and scrollbars
    tree_frame = tk.Frame(right_panel, bg="white")
    tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    # Create Treeview
    tree = ttk.Treeview(tree_frame)
    tree.grid(row=0, column=0, sticky="nsew")

    # Create Scrollbars
    scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    scrollbar_y.grid(row=0, column=1, sticky="ns")

    scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    scrollbar_x.grid(row=1, column=0, sticky="ew")

    # Configure Treeview scroll commands
    tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

    # Allow treeview frame to expand
    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)

    # Process each file
    for file in excel_files:
        try:
            file_path = os.path.join(script_dir, file)
            df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            print(f"Failed to read {file}: {e}")
            continue

        if df.empty:
            continue

        parent_id = tree.insert("", "end", text=file, open=True)

        if not tree["columns"]:
            tree["columns"] = list(df.columns)
            tree["show"] = "tree headings"

            for col in df.columns:
                tree.heading(col, text=col)
                tree.column(col, width=120, anchor="center")

        for row in df.itertuples(index=False):
            tree.insert(parent_id, "end", values=list(row))




def load_excel_file():
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if file_path:
        try:
            # Use engine to avoid auto-detection issues
            df = pd.read_excel(file_path, engine="openpyxl")

            # Get the directory where the current script is located
            script_dir = os.path.dirname(os.path.abspath(__file__))

            # Destination path in the same directory as the script
            destination_path = os.path.join(script_dir, os.path.basename(file_path))

            # Move the file
            shutil.move(file_path, destination_path)

            messagebox.showinfo("Successful", "The file has been load into the system.\nClick View all to view the data.")

            view_all()

        #     # Create new Treeview
        #     tree = ttk.Treeview(right_panel)
        #     tree.pack(expand=True, fill="both", padx=10, pady=10)
        #
        #     # Define columns
        #     tree["columns"] = list(df.columns)
        #     tree["show"] = "headings"
        #
        #     for col in df.columns:
        #         tree.heading(col, text=col)
        #         tree.column(col, width=100, anchor="center")
        #
        #     # Insert rows
        #     for _, row in df.iterrows():
        #         tree.insert("", "end", values=list(row))
        #
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")


def create_new_excel_file():
    # File name based on current date
    date_str = datetime.now().strftime("%Y-%m-%d")
    script_dir = os.path.dirname(os.path.abspath(__file__))
    file_path = os.path.join(script_dir, f"{date_str}.xlsx")

    # Data entry form
    form = tk.Toplevel(root)
    form.title("Create New Excel File")
    form.geometry("450x300")

    fields = ["Product ID", "Product Name", "Product Description", "Test Date"]
    entries = {}

    for i, field in enumerate(fields):
        tk.Label(form, text=field).grid(row=i, column=0, sticky="e", padx=10, pady=5)
        entry = tk.Entry(form, width=40)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries[field] = entry

    def save_file():
        data = {field: entry.get().strip() for field, entry in entries.items()}

        if not all(data.values()):
            messagebox.showerror("Input Error", "Please fill out all fields.")
            return

        try:
            df = pd.DataFrame([data])
            df.to_excel(file_path, index=False, engine="openpyxl")
            messagebox.showinfo("Success", f"File saved as {file_path}")
            form.destroy()
            view_all()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

    tk.Button(form, text="Save", command=save_file).grid(row=len(fields), columnspan=2, pady=15)


def add_data_to_selected_file():
    if not tree:
        messagebox.showwarning("No Selection", "Please select an Excel file node.")
        return

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select an Excel file node.")
        return

    # Get the parent node (Excel file name)
    selected_item = selected[0]
    while tree.parent(selected_item):  # Traverse up to the file node
        selected_item = tree.parent(selected_item)
    file_name = tree.item(selected_item, "text")

    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}")
        return

    # Simple form to enter new row data
    form = tk.Toplevel(root)
    form.title("Add Data")
    form.geometry("400x300")

    entries = {}
    for i, col in enumerate(df.columns):
        tk.Label(form, text=col).grid(row=i, column=0, sticky="e", pady=5, padx=5)
        entry = tk.Entry(form, width=30)
        entry.grid(row=i, column=1, pady=5, padx=5)
        entries[col] = entry

    def submit():
        new_data = {}
        for col, entry in entries.items():
            new_data[col] = entry.get()
        try:
            df.loc[len(df)] = new_data
            df.to_excel(file_path, index=False, engine="openpyxl")
            messagebox.showinfo("Success", "Data added successfully.")
            form.destroy()
            view_all()  # Refresh view
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add data:\n{e}")

    tk.Button(form, text="Submit", command=submit).grid(row=len(df.columns), columnspan=2, pady=10)


def delete_selected_data():
    if not tree:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return

    item = selected_item[0]
    parent = tree.parent(item)

    # If parent is empty, user selected the Excel filename node
    if not parent:
        messagebox.showwarning("Invalid Selection", "Please select a data row, not the file name.")
        return

    file_name = tree.item(parent, "text")
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read file:\n{e}")
        return

    # Get values from selected tree row
    row_values = tree.item(item, "values")

    # Confirm before deleting
    confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete this row?\n\n{row_values}")
    if not confirm:
        return

    # Find row to delete
    try:
        df = df[~(df.astype(str) == pd.Series(row_values, index=df.columns).astype(str)).all(axis=1)]
        df.to_excel(file_path, index=False, engine="openpyxl")
        messagebox.showinfo("Deleted", "Selected row has been deleted.")
        view_all()
    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete data:\n{e}")


def generate_barcode():
    if not tree:
        messagebox.showwarning("No Tree", "No data available.")
        return

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return

    item = selected_item[0]
    parent = tree.parent(item)

    # If parent is empty, user selected the Excel file node
    if not parent:
        messagebox.showwarning("Invalid Selection", "Please select a data row, not a file name.")
        return

    row_values = tree.item(item, "values")

    if not row_values or not row_values[0]:
        messagebox.showerror("Error", "Selected row does not have a valid ID.")
        return

    id_value = str(row_values[0])  # Assuming ID is in the first column

    try:
        script_dir = os.path.dirname(os.path.abspath(__file__))
        output_path = os.path.join(script_dir, f"barcode_{id_value}.png")

        code = barcode.get('code128', id_value, writer=ImageWriter())
        code.save(output_path)

        messagebox.showinfo("Barcode Generated", f"Barcode for ID '{id_value}' saved as:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to generate barcode:\n{e}")



# Create main window
root = tk.Tk()
# icon = tk.PhotoImage(file="transparent.png")
# root.iconphoto(False, icon)
root.title("Test")
root.geometry("1000x650")

# Create frames for left and right panels
left_panel = tk.Frame(root, bg="lightgray", width=150)
left_panel.pack(side="left", fill="y")

right_panel = tk.Frame(root, bg="white")
right_panel.pack(side="right", expand=True, fill="both")

style = ttk.Style()
style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)

# Add buttons to left panel
btn1 = ttk.Button(left_panel, text="Load File", style="Bold.TButton", command=load_excel_file)
btn1.pack(pady=(10,3), padx=8, fill="x")

btn2 = ttk.Button(left_panel, text="View all", style="Bold.TButton", command=view_all)
btn2.pack(pady=3, padx=8, fill="x")

btn3 = ttk.Button(left_panel, text="Add Batch", style="Bold.TButton", command=create_new_excel_file)
btn3.pack(pady=3, padx=8, fill="x")

# Top bar frame in right panel to hold buttons
top_bar = tk.Frame(right_panel, bg="white")
top_bar.pack(fill="x", padx=8, pady=5)

# Add a label in the right panel for content
btn4 = ttk.Button(top_bar, text="Add", style="Bold.TButton", width=15, command=add_data_to_selected_file)
btn4.pack(side="right", pady=(10,0), padx=5)

btn5 = ttk.Button(top_bar, text="Delete", style="Bold.TButton", width=15, command=delete_selected_data)
btn5.pack(side="right", pady=(10,0), padx=5)

btn6 = ttk.Button(top_bar, text="Barcode", style="Bold.TButton", width=15, command=generate_barcode)
btn6.pack(side="right", pady=(10,0), padx=5)

view_all()

# Start the main event loop
root.mainloop()
