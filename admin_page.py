import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import openpyxl
import os
import shutil
import json
from datetime import datetime, timezone
import firebase_admin
from firebase_admin import credentials, firestore

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-beb7183669.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

tree = None
tree_frame = None
top_bar = None

# # Constants
# SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# USERS_FILE = os.path.join(SCRIPT_DIR, "users.json")
# PENDING_DIR = os.path.join(SCRIPT_DIR, "pending_files")
# APPROVED_DIR = os.path.join(SCRIPT_DIR, "approved_files")
# BARCODE_REQUESTS_FILE = os.path.join(SCRIPT_DIR, "barcode_requests.csv")

# # Ensure necessary directories exist
# os.makedirs(PENDING_DIR, exist_ok=True)
# os.makedirs(APPROVED_DIR, exist_ok=True)

def approve_requests():
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


def view_pending_requests(right_panel):
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

        btn_add = ttk.Button(top_bar, text="Approve", style="Bold.TButton", width=15, command=approve_requests)
        btn_add.pack(side="right", pady=(10, 0), padx=5)

        btn_edit = ttk.Button(top_bar, text="Reject", style="Bold.TButton", width=15)
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


def admin_panel(admin_data):
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

    ttk.Button(left_panel, text="View Requests", style="Bold.TButton", command=lambda: view_pending_requests(right_panel)).pack(pady=5, padx=10)
    ttk.Button(left_panel, text="Manage Users", style="Bold.TButton").pack(pady=5, padx=10)
    ttk.Button(left_panel, text="Barcode Requests", style="Bold.TButton").pack(pady=5, padx=10)

    root.mainloop()

if __name__ == "__main__":
    dummy_user = {
        "Username": "test_user",
        "UserID": "user123",
        "Email": "test@example.com",
        "Password": "1234"
    }
    admin_panel(dummy_user)


# Helper: Clear tree

# def clear_tree():
#     global tree, tree_frame
#     if tree:
#         tree.destroy()
#     if tree_frame:
#         tree_frame.destroy()


# # Load file (goes to pending for approval)
# def load_file():
#     file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
#     if file_path:
#         dest_path = os.path.join(PENDING_DIR, f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{os.path.basename(file_path)}")
#         shutil.copy(file_path, dest_path)
#         messagebox.showinfo("Upload Complete", "File uploaded for admin approval.")

# Admin: Approve or reject pending files
# def approve_files():
#     clear_tree()
#     global tree, tree_frame
#     pending_files = [f for f in os.listdir(PENDING_DIR) if f.endswith(('.xlsx', '.xls'))]
#     if not pending_files:
#         messagebox.showinfo("No Pending Files", "No files pending approval.")
#         return
#
#     tree_frame = tk.Frame(right_panel, bg="white")
#     tree_frame.pack(expand=True, fill="both")
#     tree = ttk.Treeview(tree_frame, columns=("filename", "action"), show="headings")
#     tree.heading("filename", text="Filename")
#     tree.heading("action", text="Action")
#
#     for f in pending_files:
#         tree.insert("", "end", values=(f, "Pending"))
#     tree.pack(fill="both", expand=True)
#
#     def approve():
#         sel = tree.selection()
#         for item in sel:
#             filename = tree.item(item)['values'][0]
#             shutil.move(os.path.join(PENDING_DIR, filename), os.path.join(APPROVED_DIR, filename))
#         messagebox.showinfo("Success", "Selected files approved.")
#         approve_files()
#
#     def reject():
#         sel = tree.selection()
#         for item in sel:
#             filename = tree.item(item)['values'][0]
#             os.remove(os.path.join(PENDING_DIR, filename))
#         messagebox.showinfo("Success", "Selected files rejected.")
#         approve_files()
#
#     btn_frame = tk.Frame(tree_frame)
#     btn_frame.pack()
#     ttk.Button(btn_frame, text="Approve", style="Bold.TButton", command=approve).pack(side="left", padx=5)
#     ttk.Button(btn_frame, text="Reject", style="Bold.TButton", command=reject).pack(side="left", padx=5)
#
# # Admin: Manage users
# def manage_users():
#     clear_tree()
#     users = load_users()
#
#     def add_user():
#         uname = simpledialog.askstring("Add User", "Enter new username:")
#         pwd = simpledialog.askstring("Password", "Enter password:", show='*')
#         if uname in users:
#             messagebox.showerror("Error", "User already exists.")
#         else:
#             users[uname] = {"password": pwd, "role": "user"}
#             save_users(users)
#             messagebox.showinfo("Success", f"User {uname} added.")
#
#     def del_user():
#         uname = simpledialog.askstring("Delete User", "Enter username to delete:")
#         if uname == "admin" or uname not in users:
#             messagebox.showerror("Error", "Invalid user.")
#         else:
#             del users[uname]
#             save_users(users)
#             messagebox.showinfo("Deleted", f"User {uname} deleted.")
#
#     tk.Label(right_panel, text="User Management", font=("Segoe UI", 14)).pack(pady=10)
#     tk.Button(right_panel, text="Add User", command=add_user).pack(pady=5)
#     tk.Button(right_panel, text="Delete User", command=del_user).pack(pady=5)
#
# def approve_barcodes():
#     clear_tree()
#     if not os.path.exists(BARCODE_REQUESTS_FILE):
#         messagebox.showinfo("No Requests", "No barcode requests found.")
#         return
#     df = pd.read_csv(BARCODE_REQUESTS_FILE)
#     if df.empty:
#         messagebox.showinfo("No Requests", "No barcode requests available.")
#         return
#
#     tree_frame = tk.Frame(right_panel, bg="white")
#     tree_frame.pack(expand=True, fill="both")
#
#     tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings")
#     for col in df.columns:
#         tree.heading(col, text=col)
#         tree.column(col, width=100)
#
#     for row in df.itertuples(index=False):
#         tree.insert("", "end", values=list(row))
#     tree.pack(fill="both", expand=True)
#
#     def approve():
#         messagebox.showinfo("Approved", "(Simulated) Barcode requests approved.")
#
#     tk.Button(tree_frame, text="Approve All", command=approve).pack(pady=5)







