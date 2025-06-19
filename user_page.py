import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.simpledialog import askstring
from tkcalendar import DateEntry
import pandas as pd
import openpyxl
import os
import barcode
from barcode.writer import ImageWriter
from datetime import datetime, timezone
import firebase_admin
from firebase_admin import credentials, firestore

tree = None
tree_frame = None
top_bar = None


def insert_data_into_tree(file_name, df, right_panel):
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



def load_excel_file(right_panel, root):

    file_paths = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_paths:
        return

    load_file_layout(right_panel, root)

    for file_path in file_paths:
        file_name = os.path.basename(file_path)

        try:
            df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file:\n{file_path}\n\n{e}")
            continue

        if df.empty:
            continue

        try:
            insert_data_into_tree(file_name, df, right_panel)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to insert data:\n{e}")

    messagebox.showinfo("Done", "File(s) loaded successfully.")


def clear_right_panel(right_panel):
    global tree, tree_frame

    for widget in right_panel.winfo_children():
        widget.destroy()

    if tree:
        tree.destroy()
        tree = None

    if tree_frame:
        tree_frame.destroy()
        tree_frame = None


def create_new_excel_file(root, right_panel):
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


        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            # Create tree_frame and tree if not already created
            insert_data_into_tree(file_name, df)
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

    fields = ["Product ID", "Product Name", "Product Description", "Test date"]
    entries = {}

    for i, col in enumerate(fields):
        tk.Label(form, text=col).grid(row=i, column=0, sticky="e", padx=10, pady=5)
        if col == "Test date":
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
            if field == "Test date":
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
            insert_data_into_tree(file_name, df, right_panel)
            messagebox.showinfo("Success", f"File '{file_name}' created and saved.")
            form.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file:\n{e}")

    tk.Button(form, text="Save", command=save_file).grid(row=len(fields), columnspan=2, pady=15)


def add_data_to_selected_file(root):
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
    form.geometry("400x250")

    entries = {}
    for i, col in enumerate(df.columns):
        tk.Label(form, text=col).grid(row=i, column=0, sticky="e", pady=5, padx=5)
        if col == "Test date":
            date_entry = DateEntry(form, width=30, background='grey',
                                   foreground='white', borderwidth=1, date_pattern='dd-mm-yyyy', mindate=datetime.now())
            date_entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = date_entry
        elif col == "Product Description":
            text_desc = tk.Text(form, height=3, width=30, font=("Arial", 9))
            text_desc.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = text_desc
        else:
            entry = tk.Entry(form, width=30, font=("Arial", 9))
            entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = entry


    def submit():
        new_row = {}
        for col, entry in entries.items():
            if isinstance(entry, tk.Text):
                value = entry.get("1.0", "end-1c").strip()
            else:
                value = entry.get().strip()

            if col == "Test date":
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


def edit_selected_data(root):
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
        if col == "Test date":
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
            if col == "Test date":
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


def add_new_batch_to_pending(right_panel, root, db, user_data):
    form = tk.Toplevel(root)
    form.title("Add new Batch")
    form.geometry("400x200")

    ttk.Label(form, text="Product ID").grid(row=1, column=0, sticky="e", pady=5, padx=5)
    entry_id = ttk.Entry(form, width=30, font=("Arial", 9))
    entry_id.grid(row=1, column=1, pady=5, padx=5)

    ttk.Label(form, text="Product Name").grid(row=2, column=0, sticky="e", pady=5, padx=5)
    entry_name = ttk.Entry(form, width=30, font=("Arial", 9))
    entry_name.grid(row=2, column=1, pady=5, padx=5)

    ttk.Label(form, text="Description").grid(row=3, column=0, sticky="e", pady=5, padx=5)
    text_desc = tk.Text(form, height=3, width=30, font=("Arial", 9))
    text_desc.grid(row=3, column=1, pady=5, padx=5)

    ttk.Label(form, text="Test Date").grid(row=4, column=0, sticky="e", pady=5, padx=5)
    date_entry = DateEntry(form, width=30, background='gray',
                           foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
    date_entry.grid(row=4, column=1, pady=5, padx=5)

    def submit():
        product_id = entry_id.get().strip()
        name = entry_name.get().strip()
        desc = text_desc.get("1.0", tk.END).strip()
        test_date = date_entry.get_date()

        data = {
            "Product_ID": product_id,
            "Product_Name": name,
            "Description": desc,
            "Test_Date": test_date.strftime("%d-%m-%Y"),
            "Submitted_At": datetime.now(timezone.utc),
            "UserID": user_data.get("UserID")
        }

        try:
            db.collection("Pending").document(product_id).set(data)
            messagebox.showinfo("Success", "Product submitted for approval.")

            if tree:
                tree.insert("", "end", values=(
                    product_id,
                    name,
                    desc,
                    test_date.strftime("%d-%m-%Y"),
                    datetime.now(timezone.utc).strftime("%d-%m-%Y %H:%M")
                ))

            entry_id.delete(0, tk.END)
            entry_name.delete(0, tk.END)
            text_desc.delete("1.0", tk.END)
            date_entry.set_date(datetime.today())

            form.lift()
            form.focus_force()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")

    tk.Button(form, text="Save", command=submit).grid(row=9, columnspan=2, pady=10)


def load_pending_to_tree(right_panel, db):
    global tree, tree_frame

    if not tree_frame:
        tree_frame = tk.Frame(right_panel, bg="white")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)
    if not tree:
        tree = ttk.Treeview(tree_frame, columns=("Product ID", "Product Name", "Description", "Test Date", "Submitted At"),
            show="headings")
        tree.heading("Product ID", text="Product ID")
        tree.heading("Product Name", text="Product Name")
        tree.heading("Description", text="Description")
        tree.heading("Test Date", text="Test Date")
        tree.heading("Submitted At", text="Submitted At")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)


    try:
        # Fetch all documents from the "Pending" collection
        docs = db.collection("Pending").stream()

        for doc in docs:
            data = doc.to_dict()
            tree.insert("", "end", values=(
                data.get("Product_ID", ""),
                data.get("Product_Name", ""),
                data.get("Description", ""),
                data.get("Test_Date", ""),
                data.get("Submitted_At", "").strftime("%d-%m-%Y")
            ))
        # print([doc.to_dict() for doc in db.collection("Pending").stream()])
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load pending data:\n{e}")



def add_batch_layout(right_panel, root, db, user_data):
    global top_bar

    clear_right_panel(right_panel)

    if top_bar is not None:
        top_bar.destroy()
        top_bar = None

    if top_bar is None:
        # Top bar frame in right panel to hold buttons
        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)

        btn_add = ttk.Button(top_bar, text="Add", style="Bold.TButton", width=15, command=lambda: add_new_batch_to_pending(right_panel, root, db, user_data))
        btn_add.pack(side="right", pady=(10, 0), padx=5)

        btn_edit = ttk.Button(top_bar, text="Edit", style="Bold.TButton", width=15, command=lambda: edit_selected_data(root))
        btn_edit.pack(side="right", pady=(10, 0), padx=5)

        btn_delete = ttk.Button(top_bar, text="Delete", style="Bold.TButton", width=15, command=delete_selected_data)
        btn_delete.pack(side="right", pady=(10, 0), padx=5)

        load_pending_to_tree(right_panel, db)


def add_load_file_top_bar_buttons(root):
    global top_bar

    buttons = [
        ("Add", lambda: add_data_to_selected_file(root)),
        ("Delete", delete_selected_data),
        ("Edit", lambda: edit_selected_data(root)),
        ("Barcode", generate_barcode)
    ]

    for text, cmd in reversed(buttons):  # reversed to keep order when packing to the right
        btn = ttk.Button(top_bar, text=text, style="Bold.TButton", width=15, command=cmd)
        btn.pack(side="right", pady=(10, 0), padx=5)


def load_file_layout(right_panel, root):
    global top_bar

    clear_right_panel(right_panel)

    if top_bar is not None:
        top_bar.destroy()
        top_bar = None
    if top_bar is None:
        # Create top bar frame only once
        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)
        # Add buttons to top_bar
        add_load_file_top_bar_buttons(root)



def user_panel(user_data, db):
    # Create main window
    root = tk.Tk()
    root.title(f"Welcome, {user_data.get("Username")}")
    root.geometry("1000x650")

    # Create frames for left and right panels
    left_panel = tk.Frame(root, bg="lightgray", width=150)
    left_panel.pack(side="left", fill="y")

    right_panel = tk.Frame(root, bg="white")
    right_panel.pack(side="right", expand=True, fill="both")

    style = ttk.Style()
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)

    # Add buttons to left panel
    btn1 = ttk.Button(left_panel, text="Load File", style="Bold.TButton", command=lambda: load_excel_file(right_panel, root))
    btn1.pack(pady=(10,3), padx=8, fill="x")

    btn2 = ttk.Button(left_panel, text="Clear all", style="Bold.TButton", command=lambda: clear_right_panel(right_panel))
    btn2.pack(pady=3, padx=8, fill="x")

    btn3 = ttk.Button(left_panel, text="Add Batch", style="Bold.TButton", command=lambda: add_batch_layout(right_panel, root, db, user_data))
    btn3.pack(pady=3, padx=8, fill="x")


    # Start the main event loop
    root.mainloop()




