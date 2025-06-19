import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import openpyxl
import os
from datetime import datetime, timezone
import firebase_admin
from firebase_admin import credentials, firestore

tree = None
tree_frame = None
top_bar = None


def reject_requests(db):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row to reject.")
        return

    confirm = messagebox.askyesno("Confirm Reject", "Are you sure you want to reject the selected request(s)?")
    if not confirm:
        return

    try:
        for item in selected:
            values = tree.item(item, "values")
            product_id = values[0]

            # Delete from Firestore
            db.collection("Pending").document(product_id).delete()

            # Remove from Treeview
            tree.delete(item)

        messagebox.showinfo("Rejected", "Selected request(s) have been rejected and removed.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to reject request(s):\n{e}")


def approve_requests(db):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row to approve.")
        return

    try:
        for item in selected:
            values = tree.item(item, "values")
            product_id, name, desc, test_date, submitted_at_str, user_id = values

            # Convert submitted_at string to datetime object and get collection name
            if not submitted_at_str:
                messagebox.showerror("Error", "Missing 'Submitted At' field.")
                return

            submitted_at = datetime.strptime(submitted_at_str, "%d-%m-%Y")
            collection_name = submitted_at.strftime("%d-%m-%Y")  # Collection name and Excel file name

            # Prepare data
            data = {
                "Product_ID": product_id,
                "Product_Name": name,
                "Description": desc,
                "Test_Date": test_date,
                "Submitted_At": submitted_at.strftime("%d-%m-%Y"),
                "UserID": user_id
            }

            # Move to new collection
            db.collection(collection_name).document(product_id).set(data)
            # Delete from "Pending"
            db.collection("Pending").document(product_id).delete()

            # Write to Excel file
            file_name = f"{collection_name}.xlsx"
            file_path = os.path.join(os.getcwd(), file_name)

            data.popitem()
            df = pd.DataFrame([data])
            if os.path.exists(file_path):
                existing_df = pd.read_excel(file_path)
                df = pd.concat([existing_df, df], ignore_index=True)
            df.to_excel(file_path, index=False)

            # Remove from Treeview
            tree.delete(item)

        messagebox.showinfo("Success", "Selected request(s) approved and saved.")
    except Exception as e:
        messagebox.showerror("Error", f"Approval failed:\n{e}")


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


def create_tree_view(right_panel):
    global tree, tree_frame

    # tree view
    if not tree_frame:
        tree_frame = tk.Frame(right_panel, bg="white")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    if not tree:
        tree = ttk.Treeview(tree_frame,
                            columns=("Product ID", "Product Name", "Description", "Test Date", "Submitted At", "Product Owner"),
                            show="headings")
        tree.heading("Product ID", text="Product ID")
        tree.heading("Product Name", text="Product Name")
        tree.heading("Description", text="Description")
        tree.heading("Test Date", text="Test Date")
        tree.heading("Submitted At", text="Submitted At")
        tree.heading("Product Owner", text="Product Owner")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)


def view_pending_requests(right_panel, db):
    global top_bar

    clear_right_panel(right_panel)

    # top bar navigation
    if top_bar is not None:
        top_bar.destroy()
        top_bar = None

    if top_bar is None:
        # Top bar frame in right panel to hold buttons
        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)

        btn_add = ttk.Button(top_bar, text="Approve", style="Bold.TButton", width=15, command=lambda: approve_requests(db))
        btn_add.pack(side="right", pady=(10, 0), padx=5)

        btn_edit = ttk.Button(top_bar, text="Reject", style="Bold.TButton", width=15, command=lambda: reject_requests(db))
        btn_edit.pack(side="right", pady=(10, 0), padx=5)

    create_tree_view(right_panel)

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
                data.get("Submitted_At", "").strftime("%d-%m-%Y") if isinstance(data.get("Submitted_At"),
                                                                                 datetime) else "",
                data.get("UserID", "")
            ))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load pending data:\n{e}")


def admin_panel(admin_data, db):
    # Tkinter app setup
    root = tk.Tk()
    root.title(f"Welcome, {admin_data.get("AdminID")}")
    root.geometry("1000x650")

    left_panel = tk.Frame(root, bg="lightgray", width=180)
    left_panel.pack(side="left", fill="y")

    right_panel = tk.Frame(root, bg="white")
    right_panel.pack(side="right", expand=True, fill="both")

    style = ttk.Style()
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)

    ttk.Button(left_panel, text="View Requests", style="Bold.TButton", command=lambda: view_pending_requests(right_panel, db)).pack(pady=5, padx=10)
    ttk.Button(left_panel, text="Manage Users", style="Bold.TButton").pack(pady=5, padx=10)
    ttk.Button(left_panel, text="Barcode Requests", style="Bold.TButton").pack(pady=5, padx=10)

    root.mainloop()


