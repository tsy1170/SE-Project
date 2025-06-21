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
import Login

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

    # If columns are not set, initialize columns
    if not tree["columns"]:
        tree["columns"] = list(df.columns)
        tree["show"] = "tree headings"
        for col in df.columns:
            tree.heading(col, text=col)
            if col == "Description":
                tree.column(col, width=300, anchor="center", stretch=True)
            else:
                tree.column(col, width=150, anchor="center")


    parent_id = tree.insert("", "end", text=file_name, open=True)
    for _, row in df.iterrows():
        tree.insert(parent_id, "end", values=list(row))



def load_excel_file(right_panel, root, db):

    file_paths = filedialog.askopenfilenames(
        title="Select Excel File(s)",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_paths:
        return

    load_file_layout(right_panel, root, db)

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
        try:
            widget.destroy()
        except tk.TclError:
            pass

    if tree:
        try:
            if tree.winfo_exists():
                tree.destroy()
        except tk.TclError:
            pass
        tree = None

    if tree_frame:
        try:
            if tree_frame.winfo_exists():
                tree_frame.destroy()
        except tk.TclError:
            pass
        tree_frame = None


def delete_selected_data(db):
    if not tree:
        messagebox.showwarning("No Tree", "No data available.")
        return

    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return

    item = selected_item[0]
    parent = tree.parent(item)

    if not parent:
        messagebox.showwarning("Invalid Selection", "Please select a data row, not a file name.")
        return

    row_values = tree.item(item, "values")

    confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete this row?\n\n{row_values}")
    if not confirm:
        return

    try:
        product_id = row_values[0]
        submitted_at = row_values[-1]  # Get the last column (Submitted_At)

        # Construct Excel filename and path
        file_name = f"{submitted_at}.xlsx"
        file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

        # Load the Excel file
        df = pd.read_excel(file_path, engine="openpyxl")

        # Remove the row by matching Product_ID
        df = df[df["Product_ID"].astype(str) != str(product_id)]

        # Save the updated Excel file
        df.to_excel(file_path, index=False, engine="openpyxl")

        # Delete from Firestore
        db.collection(submitted_at).document(product_id).delete()

        # Remove the item from the Treeview
        tree.delete(item)

        messagebox.showinfo("Deleted", f"Product '{product_id}' has been deleted.")

    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete data:\n{e}")


def edit_selected_data(root, db):
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

    # Get file and path
    file_name = tree.item(parent, "text")  # e.g., "12-07-2024"
    file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), file_name)

    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read Excel:\n{e}")
        return

    row_values = tree.item(item, "values")
    original_index = None

    for i, row in df.iterrows():
        if all(str(row[col]) == str(row_values[j]) for j, col in enumerate(df.columns)):
            original_index = i
            break

    if original_index is None:
        messagebox.showerror("Error", "Row not found in Excel.")
        return

    # ---------------- Edit UI ----------------
    form = tk.Toplevel(root)
    form.configure(bg="White")
    form.title("Edit Product")
    form.geometry("450x350")
    entries = {}

    for i, col in enumerate(df.columns):
        ttk.Label(form, text=col+":", style="Custom.TLabel").grid(row=i, column=0, sticky="e", pady=5, padx=(37, 10))

        if col == "Product_ID" or col == "Submitted_At":
            entry = ttk.Entry(form, width=30, font=("Arial", 9))
            entry.insert(0, row_values[i])
            entry.configure(state="disabled")  # Make it read-only
            entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = entry  # Store the widget, not just the value

        elif col == "Test_Date":
            date_entry = DateEntry(
                form, width=28, background='grey',
                foreground='white', borderwidth=1, date_pattern='dd-mm-yyyy',
                mindate=datetime.today()
            )
            date_entry.set_date(datetime.strptime(row_values[i], "%d-%m-%Y"))
            date_entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = date_entry

        elif col == "Description":
            entry = tk.Text(form, height=3, width=30, font=("Arial", 9))
            entry.configure(highlightthickness=1, highlightbackground="#ccc", highlightcolor="#4a90e2")
            entry.insert("1.0", row_values[i])
            entry.grid(row=2, column=1, pady=5, padx=5)
            entries[col] = entry

        else:
            entry = ttk.Entry(form, width=30, font=("Arial", 9))
            entry.insert(0, row_values[i])
            entry.grid(row=i, column=1, pady=5, padx=5)
            entries[col] = entry

    # --------------- Save Changes ----------------
    def save_edit():
        updated_row = {}

        for col in df.columns:
            widget = entries[col]

            if col in ["Product_ID", "Submitted_At"]:
                updated_row[col] = widget.get()
            elif col == "Test_Date":
                date_obj = widget.get_date()
                if date_obj < datetime.today().date():
                    messagebox.showerror("Date Error", "Test Date cannot be in the past.")
                    return
                updated_row[col] = date_obj.strftime("%d-%m-%Y")
            else:
                value = widget.get().strip()
                if not value:
                    messagebox.showerror("Input Error", f"{col} cannot be empty.")
                    return
                updated_row[col] = value

        try:
            # Update DataFrame
            df.loc[original_index] = updated_row
            df.to_excel(file_path, index=False, engine="openpyxl")

            # Update Firestore
            product_id = updated_row["Product_ID"]
            collection_name = updated_row["Submitted_At"]
            db.collection(collection_name).document(product_id).update(updated_row)

            # Refresh Treeview
            for child in tree.get_children(parent):
                tree.delete(child)
            for _, row in df.iterrows():
                tree.insert(parent, "end", values=list(row))

            messagebox.showinfo("Success", "Data updated.")
            form.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update:\n{e}")

    ttk.Button(form, text="Save Changes", command=save_edit).grid(row=len(df.columns), columnspan=2, pady=15, padx=(60, 0))


def edit_pending_items(root, db):
    global tree

    selected = tree.selection()
    if len(selected) != 1:
        messagebox.showwarning("Selection Error", "Please select exactly one item to edit.")
        return

    item_id = selected[0]
    values = tree.item(item_id, "values")
    product_id, name, desc, test_date, submitted_at_str = values[:5]

    # Create Edit Form
    form = tk.Toplevel(root)
    form.configure(bg="White")
    form.title("Edit Item")
    form.geometry("420x270")

    ttk.Label(form, text="Product ID:", style="Custom.TLabel").grid(row=0, column=0, sticky="e", pady=5, padx=(24, 10))
    entry_id = ttk.Entry(form, width=30, font=("Arial", 9))
    entry_id.insert(0, product_id)
    entry_id.config(state="disabled")  # Product ID should not be editable
    entry_id.grid(row=0, column=1, pady=5, padx=5)

    ttk.Label(form, text="Product Name:", style="Custom.TLabel").grid(row=1, column=0, sticky="e", pady=5, padx=(24, 10))
    entry_name = ttk.Entry(form, width=30, font=("Arial", 9))
    entry_name.insert(0, name)
    entry_name.grid(row=1, column=1, pady=5, padx=5)

    ttk.Label(form, text="Description:", style="Custom.TLabel").grid(row=2, column=0, sticky="e", pady=5, padx=(24, 10))
    text_desc = tk.Text(form, height=3, width=30, font=("Arial", 9))
    text_desc.configure(highlightthickness=1, highlightbackground="#ccc", highlightcolor="#4a90e2")
    text_desc.insert("1.0", desc)
    text_desc.grid(row=2, column=1, pady=5, padx=5)

    ttk.Label(form, text="Test Date:", style="Custom.TLabel").grid(row=3, column=0, sticky="e", pady=5, padx=(24, 10))
    date_entry = DateEntry(form, width=30, background='gray',
                           foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
    date_entry.set_date(datetime.strptime(test_date, "%d-%m-%Y"))
    date_entry.grid(row=3, column=1, pady=5, padx=5)

    def update():
        new_name = entry_name.get().strip()
        new_desc = text_desc.get("1.0", tk.END).strip()
        new_test_date = date_entry.get_date()

        if not new_name or not new_desc:
            messagebox.showwarning("Input Error", "All fields must be filled.")
            form.lift()
            form.focus_force()
            return

        if new_test_date < datetime.today().date():
            messagebox.showwarning("Invalid Date", "Test date cannot be in the past.")
            form.lift()
            form.focus_force()
            return

        try:
            doc_ref = db.collection("Pending").document(product_id)
            doc_data = doc_ref.get().to_dict()

            doc_data.update({
                "Product_Name": new_name,
                "Description": new_desc,
                "Test_Date": new_test_date.strftime("%d-%m-%Y")
            })

            doc_ref.set(doc_data)

            # Update treeview
            tree.item(item_id, values=(
                product_id,
                new_name,
                new_desc,
                new_test_date.strftime("%d-%m-%Y"),
                submitted_at_str
            ))

            form.destroy()
            messagebox.showinfo("Success", "Item updated successfully.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to update item:\n{e}")

    ttk.Button(form, text="Update", command=update).grid(row=5, columnspan=2, pady=10, padx=(45, 0))



def delete_pending_items(db):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select item(s) to delete.")
        return

    confirm = messagebox.askyesno("Confirm Deletion", "Are you sure you want to delete the selected item(s)?")
    if not confirm:
        return

    try:
        for item in selected:
            values = tree.item(item, "values")
            product_id = values[0]  # Assuming 'Product ID' is the first column

            # Delete from Firestore "Pending" collection
            db.collection("Pending").document(product_id).delete()

            # Remove from Treeview
            tree.delete(item)

        messagebox.showinfo("Success", "Selected item(s) deleted from Pending.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete item(s):\n{e}")



def add_items_to_pending(right_panel, root, db, user_data):
    form = tk.Toplevel(root)
    form.configure(bg="White")
    form.title("Add new Batch")
    form.geometry("420x270")

    ttk.Label(form, text="Product ID:", style="Custom.TLabel").grid(row=1, column=0, sticky="e", pady=5, padx=(24, 10))
    entry_id = ttk.Entry(form, width=30, font=("Arial", 9))
    entry_id.grid(row=1, column=1, pady=5, padx=5)

    ttk.Label(form, text="Product Name:", style="Custom.TLabel").grid(row=2, column=0, sticky="e", pady=5, padx=(24, 10))
    entry_name = ttk.Entry(form, width=30, font=("Arial", 9))
    entry_name.grid(row=2, column=1, pady=5, padx=5)

    ttk.Label(form, text="Description:", style="Custom.TLabel").grid(row=3, column=0, sticky="e", pady=5, padx=(24, 10))
    text_desc = tk.Text(form, height=3, width=30, font=("Arial", 9))
    text_desc.configure(highlightthickness=1, highlightbackground="#ccc", highlightcolor="#4a90e2")
    text_desc.grid(row=3, column=1, pady=5, padx=5)

    ttk.Label(form, text="Test Date:", style="Custom.TLabel").grid(row=4, column=0, sticky="e", pady=5, padx=(24, 10))
    date_entry = DateEntry(form, width=30, background='gray',
                           foreground='white', borderwidth=2, date_pattern='dd-mm-yyyy')
    date_entry.grid(row=4, column=1, pady=5, padx=5)

    def submit():
        product_id = entry_id.get().strip()
        name = entry_name.get().strip()
        desc = text_desc.get("1.0", tk.END).strip()
        test_date = date_entry.get_date()
        today = datetime.today().date()

        # Check for missing fields
        if not product_id or not name or not desc:
            messagebox.showwarning("Input Error", "Please fill in all fields.")
            form.lift()
            form.focus_force()
            return

        # Check if Test Date is not before today
        if test_date < today:
            messagebox.showwarning("Invalid Date", "Test Date cannot be before today.")
            form.lift()
            form.focus_force()
            return

        # Check for duplicate Product ID
        doc_ref = db.collection("Pending").document(product_id)
        if doc_ref.get().exists:
            messagebox.showerror("Duplicate Product ID", f"Product ID '{product_id}' already exists in Pending.")
            form.lift()
            form.focus_force()
            return

        try:
            data = {
                "Product_ID": product_id,
                "Product_Name": name,
                "Description": desc,
                "Test_Date": test_date.strftime("%d-%m-%Y"),
                "Submitted_At": datetime.now(timezone.utc),
                "UserID": user_data.get("UserID")
            }


            db.collection("Pending").document(product_id).set(data)
            messagebox.showinfo("Success", "Product submitted for approval.")

            if tree:
                tree.insert("", "end", values=(
                    product_id,
                    name,
                    desc,
                    test_date.strftime("%d-%m-%Y"),
                    datetime.now(timezone.utc).strftime("%d-%m-%Y")
                ))

            entry_id.delete(0, tk.END)
            entry_name.delete(0, tk.END)
            text_desc.delete("1.0", tk.END)
            date_entry.set_date(datetime.today())

            form.lift()
            form.focus_force()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to save: {e}")


    ttk.Button(form, text="Save", command=submit).grid(row=9, columnspan=2, pady=10, padx=(45, 0))


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
        try:
            if top_bar.winfo_exists():
                top_bar.destroy()
        except tk.TclError:
            pass
        top_bar = None

    if top_bar is None:
        # Top bar frame in right panel to hold buttons
        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)

        btn_add = ttk.Button(top_bar, text="Add", style="Bold.TButton", width=15, command=lambda: add_items_to_pending(right_panel, root, db, user_data))
        btn_add.pack(side="right", pady=(10, 0), padx=5)

        btn_edit = ttk.Button(top_bar, text="Delete", style="Bold.TButton", width=15, command=lambda: delete_pending_items(db))
        btn_edit.pack(side="right", pady=(10, 0), padx=5)

        btn_delete = ttk.Button(top_bar, text="Edit", style="Bold.TButton", width=15, command=lambda: edit_pending_items(root, db))
        btn_delete.pack(side="right", pady=(10, 0), padx=5)

        load_pending_to_tree(right_panel, db)


def load_file_layout(right_panel, root, db):
    global top_bar

    clear_right_panel(right_panel)

    if top_bar is not None:
        try:
            if top_bar.winfo_exists():
                top_bar.destroy()
        except tk.TclError:
            pass  # The widget is already invalid or belongs to an old root
        top_bar = None

    if top_bar is None:
        # Create top bar frame only once
        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)

        # Add buttons to top_bar
        btn_delete = ttk.Button(top_bar, text="Delete", style="Bold.TButton", width=15, command=lambda: delete_selected_data(db))
        btn_delete.pack(side="right", pady=(10, 0), padx=5)
        btn_edit = ttk.Button(top_bar, text="Edit", style="Bold.TButton", width=15, command=lambda: edit_selected_data(root, db))
        btn_edit.pack(side="right", pady=(10, 0), padx=5)

def logout(root):
    global top_bar
    top_bar = None
    root.destroy()
    Login.show_login()


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
    style.theme_use("clam")
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)
    style.configure("Treeview.Heading", background="#d3d3d3", foreground="black", font=("Segoe UI", 10, "bold"))
    style.configure("Custom.TLabel", background="White", foreground="#333333", font=("Segoe UI", 10, "bold"), padding=5)

    # Add buttons to left panel
    btn_load_file = ttk.Button(left_panel, text="Load File", style="Bold.TButton", command=lambda: load_excel_file(right_panel, root, db))
    btn_load_file.pack(pady=(10,3), padx=8, fill="x")

    btn_clear = ttk.Button(left_panel, text="Clear all", style="Bold.TButton", command=lambda: clear_right_panel(right_panel))
    btn_clear.pack(pady=3, padx=8, fill="x")

    btn_add_batch = ttk.Button(left_panel, text="Add Batch", style="Bold.TButton", command=lambda: add_batch_layout(right_panel, root, db, user_data))
    btn_add_batch.pack(pady=3, padx=8, fill="x")

    ttk.Button(left_panel, text="Logout", style="Bold.TButton", command=lambda: logout(root)).pack(pady=20, padx=10, side="bottom")

    # Start the main event loop
    root.mainloop()
