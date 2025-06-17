import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.simpledialog import askstring
from tkcalendar import DateEntry
import pandas as pd
import openpyxl
import os
import shutil
import barcode
from barcode.writer import ImageWriter
from datetime import datetime

tree = None
tree_frame = None
loaded_files = set()


def insert_data_into_tree(file_name, df):
    global tree, tree_frame

    # Create tree_frame and tree if not already created
    if not tree_frame:
        tree_frame = tk.Frame(right_panel, bg="white")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    if not tree:
        tree = ttk.Treeview(tree_frame)
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # If columns are not set, initialize them
    if not tree["columns"]:
        tree["columns"] = list(df.columns)
        tree["show"] = "tree headings"
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")

    # If columns are already set, check for mismatch
    elif list(df.columns) != list(tree["columns"]):
        messagebox.showwarning("Column Mismatch", f"File '{file_name}' skipped due to column mismatch.")
        return

    parent_id = tree.insert("", "end", text=file_name, open=True)
    for _, row in df.iterrows():
        tree.insert(parent_id, "end", values=list(row))



def load_excel_file():
    global loaded_files

    file_paths = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_paths:
        return

    script_dir = os.path.dirname(os.path.abspath(__file__))

    for file_path in file_paths:
        file_name = os.path.basename(file_path)

        # Skip if already loaded
        if file_name in loaded_files:
            messagebox.showinfo("Skipped", f"'{file_name}' is already loaded.")
            continue

        try:
            df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{file_path}\n\n{e}")
            continue

        if df.empty:
            continue

        try:
            insert_data_into_tree(file_name, df)
            loaded_files.add(file_name)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to insert data:\n{e}")

    messagebox.showinfo("Done", "File(s) loaded successfully.")


def clear_all():
    global tree, tree_frame, loaded_files

    if tree:
        tree.destroy()
        tree = None

    if tree_frame:
        tree_frame.destroy()
        tree_frame = None

    loaded_files.clear()


def create_new_excel_file():
    global tree, tree_frame

    script_dir = os.path.dirname(os.path.abspath(__file__))
    date_str = datetime.now().strftime("%Y-%m-%d")
    default_file_name = f"{date_str}.xlsx"
    default_file_path = os.path.join(script_dir, default_file_name)

    use_default = messagebox.askyesno("Create New File", f"Use default file name: {default_file_name}?")

    if use_default:
        file_path = default_file_path
        file_name = default_file_name
    else:
        user_input = askstring("File Name", "Enter a name for the new Excel file (without extension):")
        if not user_input:
            return  # Cancelled
        file_name = f"{user_input}.xlsx"
        file_path = os.path.join(script_dir, file_name)

    # If file exists, just load it
    if os.path.exists(file_path):
        if file_name in loaded_files:
            messagebox.showinfo("Already Loaded", f"The file '{file_name}' is already loaded and displayed.")
            return

        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            # Create tree_frame and tree if not already created
            insert_data_into_tree(file_name, df)
            loaded_files.add(file_name)
            if df.empty:
                messagebox.showinfo("Empty File", f"'{file_name}' is empty. You can add data.")
            return
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")
            return

    # File doesn't exist - create new and prompt for data entry
    form = tk.Toplevel(root)
    form.title("Add data")
    form.geometry("450x200")

    fields = ["Product ID", "Product Name", "Product Description", "Maturation date"]
    entries = {}

    for i, col in enumerate(fields):
        tk.Label(form, text=col).grid(row=i, column=0, sticky="e", padx=10, pady=5)
        if col == "Maturation date":
            date_entry = DateEntry(form, width=28, background='grey',
                                   foreground='white', borderwidth=1, date_pattern='dd-mm-yyyy', mindate=datetime.now())
            date_entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = date_entry
        else:
            entry = tk.Entry(form, width=30)
            entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = entry

    def save_file():
        data = {}
        for field, entry in entries.items():
            value = entry.get().strip()
            if field == "Maturation date":
                try:
                    selected_date = datetime.strptime(value, "%d-%m-%Y")
                    if selected_date.date() < datetime.today().date():
                        raise ValueError("Date must not be in the past.")
                    data[field] = selected_date.strftime("%d-%m-%Y")
                except Exception as e:
                    messagebox.showerror("Date Error", f"Invalid date: {e}")
                    return
            else:
                data[field] = value

        if not all(data.values()):
            messagebox.showerror("Input Error", "Please fill out all fields.")
            return

        try:
            df = pd.DataFrame([data])
            df.to_excel(file_path, index=False, engine="openpyxl")
            loaded_files.add(file_name)
            insert_data_into_tree(file_name, df)
            messagebox.showinfo("Success", f"File '{file_name}' created and saved.")
            form.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

    tk.Button(form, text="Save", command=save_file).grid(row=len(fields), columnspan=2, pady=15)


def add_data_to_selected_file():
    if not tree:
        messagebox.showwarning("No Tree", "No data available.")
        return

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a file or row.")
        return

    # Find the file node (root of selected branch)
    selected_item = selected[0]
    while tree.parent(selected_item):
        selected_item = tree.parent(selected_item)

    file_name = tree.item(selected_item, "text")
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}")
        return

    form = tk.Toplevel(root)
    form.title(f"Add Data to {file_name}")
    form.geometry("400x200")

    entries = {}
    for i, col in enumerate(df.columns):
        tk.Label(form, text=col).grid(row=i, column=0, sticky="e", pady=5, padx=5)
        if col == "Maturation date":
            date_entry = DateEntry(form, width=28, background='grey',
                                   foreground='white', borderwidth=1, date_pattern='dd-mm-yyyy', mindate=datetime.now())
            date_entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = date_entry
        else:
            entry = tk.Entry(form, width=30)
            entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = entry

    def submit():
        new_row = {}
        for col, entry in entries.items():
            value = entry.get().strip()
            if col == "Maturation date":
                try:
                    selected_date = datetime.strptime(value, "%d-%m-%Y")
                    if selected_date.date() < datetime.today().date():
                        raise ValueError("Date must not be in the past.")
                    new_row[col] = selected_date.strftime("%d-%m-%Y")
                except Exception as e:
                    messagebox.showerror("Date Error", f"Invalid date: {e}")
                    return
            else:
                new_row[col] = str(value)

        if not all(new_row.values()):
            messagebox.showerror("Input Error", "Please fill in all fields.")
            return

        try:
            df.loc[len(df)] = new_row
            df.to_excel(file_path, index=False, engine="openpyxl")

            for child in tree.get_children(selected_item):
                tree.delete(child)
            for _, row in df.iterrows():
                tree.insert(selected_item, "end", values=list(row))

            messagebox.showinfo("Success", "Data added and displayed.")
            form.destroy()
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

    row_values = tree.item(item, "values")

    confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete this row?\n\n{row_values}")
    if not confirm:
        return

    try:
        # Filter out the matching row from DataFrame
        df = df[~(df.astype(str) == pd.Series(row_values, index=df.columns).astype(str)).all(axis=1)]

        # Save updated data
        df.to_excel(file_path, index=False, engine="openpyxl")

        # Remove item from Treeview
        tree.delete(item)

        messagebox.showinfo("Deleted", "Selected row has been deleted.")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete data:\n{e}")


def edit_selected_data():
    if not tree:
        messagebox.showwarning("No Tree", "No data available.")
        return

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a data row to edit.")
        return

    item = selected_item[0]
    parent = tree.parent(item)

    if not parent:
        messagebox.showwarning("Invalid Selection", "Please select a data row, not a file name.")
        return

    file_name = tree.item(parent, "text")
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}")
        return

    row_values = tree.item(item, "values")
    original_index = None

    # Find the index of the selected row in the DataFrame
    for i, row in df.iterrows():
        if all(str(row[col]) == str(row_values[j]) for j, col in enumerate(df.columns)):
            original_index = i
            break

    if original_index is None:
        messagebox.showerror("Error", "Could not find the selected row in the file.")
        return

    form = tk.Toplevel(root)
    form.title(f"Edit Data in {file_name}")
    form.geometry("400x200")

    entries = {}
    for i, col in enumerate(df.columns):
        tk.Label(form, text=col).grid(row=i, column=0, sticky="e", pady=5, padx=5)
        if col == "Maturation date":
            date_entry = DateEntry(form, width=28, background='grey',
                                   foreground='white', borderwidth=1, date_pattern='dd-mm-yyyy', mindate=datetime.now())
            date_entry.set_date(datetime.strptime(row_values[i], "%d-%m-%Y"))
            date_entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = date_entry
        else:
            entry = tk.Entry(form, width=30)
            entry.insert(0, row_values[i])
            entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = entry

    def save_edit():
        updated_row = {}
        for col, entry in entries.items():
            value = entry.get().strip()
            if col == "Maturation date":
                try:
                    selected_date = datetime.strptime(value, "%d-%m-%Y")
                    if selected_date.date() < datetime.today().date():
                        raise ValueError("Date must not be in the past.")
                    updated_row[col] = selected_date.strftime("%d-%m-%Y")
                except Exception as e:
                    messagebox.showerror("Date Error", f"Invalid date: {e}")
                    return
            else:
                updated_row[col] = str(value)

        if not all(updated_row.values()):
            messagebox.showerror("Input Error", "Please fill in all fields.")
            return

        try:
            # Update the DataFrame and save
            df.loc[original_index] = updated_row
            df.to_excel(file_path, index=False, engine="openpyxl")

            # Refresh Treeview
            for child in tree.get_children(parent):
                tree.delete(child)
            for _, row in df.iterrows():
                tree.insert(parent, "end", values=list(row))

            messagebox.showinfo("Updated", "Data updated successfully.")
            form.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update data:\n{e}")

    tk.Button(form, text="Save", command=save_edit).grid(row=len(df.columns), columnspan=2, pady=10)



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

btn2 = ttk.Button(left_panel, text="Clear all", style="Bold.TButton", command=clear_all)
btn2.pack(pady=3, padx=8, fill="x")

btn3 = ttk.Button(left_panel, text="Add Files", style="Bold.TButton", command=create_new_excel_file)
btn3.pack(pady=3, padx=8, fill="x")

# Top bar frame in right panel to hold buttons
top_bar = tk.Frame(right_panel, bg="white")
top_bar.pack(fill="x", padx=8, pady=5)

# Add a label in the right panel for content
btn4 = ttk.Button(top_bar, text="Add", style="Bold.TButton", width=15, command=add_data_to_selected_file)
btn4.pack(side="right", pady=(10,0), padx=5)

btn5 = ttk.Button(top_bar, text="Delete", style="Bold.TButton", width=15, command=delete_selected_data)
btn5.pack(side="right", pady=(10,0), padx=5)

btn6 = ttk.Button(top_bar, text="Edit", style="Bold.TButton", width=15, command=edit_selected_data)
btn6.pack(side="right", pady=(10,0), padx=5)

btn7 = ttk.Button(top_bar, text="Barcode", style="Bold.TButton", width=15, command=generate_barcode)
btn7.pack(side="right", pady=(10,0), padx=5)

# Start the main event loop
root.mainloop()
