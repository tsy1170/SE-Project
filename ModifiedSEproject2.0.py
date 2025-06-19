import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import openpyxl
import os
import shutil
import json
from datetime import datetime
from admin_function import view_requests, generate_barcode_requests, init_requests, load_requests, save_requests
import yagmail

# Set script directory and init files
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
init_requests(SCRIPT_DIR)
USERS_FILE = os.path.join(SCRIPT_DIR, "users.json")
REQUESTS_FILE = os.path.join(SCRIPT_DIR, "requests.json")

def load_users():
    if not os.path.exists(USERS_FILE):
        with open(USERS_FILE, "w") as f:
            json.dump({"admin": {"password": "admin", "role": "admin"}}, f)
    with open(USERS_FILE, "r") as f:
        return json.load(f)

def save_users(users):
    with open(USERS_FILE, "w") as f:
        json.dump(users, f, indent=2)


# Login system
if not os.path.exists(USERS_FILE):
    with open(USERS_FILE, "w") as f:
        json.dump({"admin": {"password": "admin", "role": "admin"}}, f)

with open(USERS_FILE, "r") as f:
    users_data = json.load(f)

username = simpledialog.askstring("Login", "Enter username:")
password = simpledialog.askstring("Login", "Enter password:", show='*')

if username not in users_data or users_data[username]["password"] != password:
    messagebox.showerror("Login Failed", "Invalid username or password.")
    exit()

role = users_data[username]["role"]
is_admin = role == "admin"

tree_ref = [None]
tree_frame_ref = [None]

def clear_tree(tree_ref, tree_frame_ref):
    if tree_ref[0]:
        tree_ref[0].destroy()
        tree_ref[0] = None
    if tree_frame_ref[0]:
        tree_frame_ref[0].destroy()
        tree_frame_ref[0] = None

def submit_row_request(action, file_name, row_data):
    requests = load_requests(SCRIPT_DIR)
    new_request = {
        "username": username,
        "action": action,
        "file": file_name,
        "data": row_data,
        "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "status": "pending"
    }
    requests["requests"].append(new_request)
    save_requests(SCRIPT_DIR, requests)
    messagebox.showinfo("Submitted", f"Your {action} request has been sent to the admin.")

def view_my_requests():
    clear_tree(tree_ref, tree_frame_ref)
    all_data = load_requests(SCRIPT_DIR)["requests"]
    my_data = [(i, r) for i, r in enumerate(all_data) if r["username"] == username]

    if not my_data:
        messagebox.showinfo("No Requests", "You have no submitted requests.")
        return

    tree_frame_ref[0] = tk.Frame(right_panel, bg="white")
    tree_frame_ref[0].pack(expand=True, fill="both", padx=10, pady=10)

    columns = ["index", "action", "file", "status", "timestamp", "data"]
    tree = ttk.Treeview(tree_frame_ref[0], columns=columns, show="headings")
    tree_ref[0] = tree

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")

    for i, req in my_data:
        tree.insert("", "end", iid=i, values=(i, req["action"], req["file"], req["status"], req["timestamp"], str(req["data"])))

    tree.pack(expand=True, fill="both")

    def cancel_request():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a request to cancel.")
            return
        req_id = int(sel[0])
        if all_data[req_id]["status"] != "pending":
            messagebox.showwarning("Not Allowed", "Only pending requests can be cancelled.")
            return
        confirm = messagebox.askyesno("Confirm", "Cancel this request?")
        if confirm:
            all_data.pop(req_id)
            save_requests(SCRIPT_DIR, {"requests": all_data})
            messagebox.showinfo("Cancelled", "Request has been cancelled.")
            view_my_requests()

    tk.Button(tree_frame_ref[0], text="Cancel Selected Request", command=cancel_request).pack(pady=5)

def add_data_to_selected_file():
    tree = tree_ref[0]
    if not tree:
        messagebox.showwarning("No Selection", "Please select an Excel file node.")
        return
    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select an Excel file node.")
        return
    selected_item = selected[0]
    while tree.parent(selected_item):
        selected_item = tree.parent(selected_item)
    file_name = tree.item(selected_item, "text")
    file_path = os.path.join(SCRIPT_DIR, file_name)
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to read file:\n{e}")
        return

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
        new_data = {col: entry.get() for col, entry in entries.items()}
        if is_admin:
            try:
                df.loc[len(df)] = new_data
                df.to_excel(file_path, index=False, engine="openpyxl")
                messagebox.showinfo("Success", "Data added successfully.")
                form.destroy()
                view_all()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to add data:\n{e}")
        else:
            submit_row_request("add", file_name, new_data)
            form.destroy()

    tk.Button(form, text="Submit", command=submit).grid(row=len(df.columns), columnspan=2, pady=10)

def delete_selected_data():
    tree = tree_ref[0]
    if not tree:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return
    item = selected_item[0]
    parent = tree.parent(item)
    if not parent:
        messagebox.showwarning("Invalid Selection", "Please select a data row, not the file name.")
        return
    file_name = tree.item(parent, "text")
    file_path = os.path.join(SCRIPT_DIR, file_name)
    try:
        df = pd.read_excel(file_path, engine="openpyxl")
    except Exception as e:
        messagebox.showerror("Error", f"Could not read file:\n{e}")
        return
    row_values = tree.item(item, "values")
    confirm = messagebox.askyesno("Confirm Deletion", f"Delete this row?\n\n{row_values}")
    if not confirm:
        return
    row_data = dict(zip(df.columns, row_values))
    if is_admin:
        try:
            df = df[~(df.astype(str) == pd.Series(row_values, index=df.columns).astype(str)).all(axis=1)]
            df.to_excel(file_path, index=False, engine="openpyxl")
            messagebox.showinfo("Deleted", "Selected row has been deleted.")
            view_all()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to delete data:\n{e}")
    else:
        submit_row_request("delete", file_name, row_data)

def view_all():
    clear_tree(tree_ref, tree_frame_ref)
    excel_files = [f for f in os.listdir(SCRIPT_DIR) if f.endswith((".xlsx", ".xls"))]
    if not excel_files:
        messagebox.showerror("Error", "No Excel files found.")
        return
    tree_frame_ref[0] = tk.Frame(right_panel, bg="white")
    tree_frame_ref[0].pack(expand=True, fill="both", padx=10, pady=10)
    tree_ref[0] = ttk.Treeview(tree_frame_ref[0])
    tree_ref[0].grid(row=0, column=0, sticky="nsew")
    scrollbar_y = ttk.Scrollbar(tree_frame_ref[0], orient="vertical", command=tree_ref[0].yview)
    scrollbar_y.grid(row=0, column=1, sticky="ns")
    scrollbar_x = ttk.Scrollbar(tree_frame_ref[0], orient="horizontal", command=tree_ref[0].xview)
    scrollbar_x.grid(row=1, column=0, sticky="ew")
    tree_ref[0].configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
    tree_frame_ref[0].grid_rowconfigure(0, weight=1)
    tree_frame_ref[0].grid_columnconfigure(0, weight=1)
    for file in excel_files:
        try:
            file_path = os.path.join(SCRIPT_DIR, file)
            df = pd.read_excel(file_path, engine="openpyxl")
        except Exception as e:
            print(f"Failed to read {file}: {e}")
            continue
        if df.empty:
            continue
        parent_id = tree_ref[0].insert("", "end", text=file, open=True)
        if not tree_ref[0]["columns"]:
            tree_ref[0]["columns"] = list(df.columns)
            tree_ref[0]["show"] = "tree headings"
            for col in df.columns:
                tree_ref[0].heading(col, text=col)
                tree_ref[0].column(col, width=120, anchor="center")
        for row in df.itertuples(index=False):
            tree_ref[0].insert(parent_id, "end", values=list(row))

def load_excel_file():
    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if file_path:
        try:
            df = pd.read_excel(file_path, engine="openpyxl")
            destination_path = os.path.join(SCRIPT_DIR, os.path.basename(file_path))
            shutil.move(file_path, destination_path)
            messagebox.showinfo("Successful", "File loaded successfully. Click 'View all' to refresh.")
            view_all()
        except Exception as e:
            messagebox.showerror("Error", f"Could not load file:\n{e}")
def request_barcode_from_selection():
    tree = tree_ref[0]
    if not tree:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return
    selected_item = tree.selection()
    if not selected_item:
        messagebox.showwarning("No Selection", "Please select a data row.")
        return
    item = selected_item[0]
    parent = tree.parent(item)
    if not parent:
        messagebox.showwarning("Invalid Selection", "Please select a data row, not the file name.")
        return
    file_name = tree.item(parent, "text")
    row_values = tree.item(item, "values")
    row_data = dict(zip(tree["columns"], row_values))
    submit_row_request("barcode", file_name, row_data)

def open_file_by_barcode(barcode_data):
    try:
        filename, row_info = barcode_data.split("::")
        row_dict = json.loads(row_info)

        # Get the test date and convert to folder name
        date_str = row_dict.get("Test Date") or row_dict.get("Date")
        if not date_str:
            raise ValueError("No date found in barcode.")

        folder_name = datetime.strptime(date_str, "%Y-%m-%d").strftime("%Y%m%d")
        full_path = os.path.join(SCRIPT_DIR, "saved_by_date", folder_name, filename)

        if not os.path.exists(full_path):
            messagebox.showerror("File Not Found", f"File not found:\n{full_path}")
            return

        df = pd.read_excel(full_path, engine="openpyxl")

        # Display in treeview
        clear_tree(tree_ref, tree_frame_ref)
        tree_frame_ref[0] = tk.Frame(right_panel, bg="white")
        tree_frame_ref[0].pack(expand=True, fill="both", padx=10, pady=10)

        tree = ttk.Treeview(tree_frame_ref[0], columns=list(df.columns), show="headings")
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")

        for row in df.itertuples(index=False):
            tree.insert("", "end", values=list(row))

        tree.pack(expand=True, fill="both")
        tree_ref[0] = tree
        messagebox.showinfo("Success", f"File opened from folder: {folder_name}")

    except Exception as e:
        messagebox.showerror("Barcode Error", f"Failed to open file:\n{e}")

def test_open_barcode():
    barcode_input = simpledialog.askstring("Barcode Input", "Paste barcode content:")
    if barcode_input:
        open_file_by_barcode(barcode_input)


def send_email_reminders():
    try:
        with open(os.path.join(SCRIPT_DIR, "email_credentials.json")) as f:
            creds = json.load(f)
        yag = yagmail.SMTP(creds["email"], creds["password"])
    except Exception as e:
        messagebox.showerror("Email Setup Error", f"Could not set up email client:\n{e}")
        return

    files = [f for f in os.listdir(SCRIPT_DIR) if f.endswith(('.xlsx', '.xls'))]
    today = datetime.today().date()

    for file in files:
        try:
            df = pd.read_excel(os.path.join(SCRIPT_DIR, file), engine="openpyxl")
            if "Test Date" not in df.columns or "Owner Email" not in df.columns:
                continue

            for _, row in df.iterrows():
                test_date = pd.to_datetime(row["Test Date"]).date()
                if test_date == today + pd.DateOffset(months=2):
                    owner_email = row["Owner Email"]
                    product = row.get("Product", "Unnamed Product")

                    subject = "Reminder: Sample Maturation Approaching"
                    body = f"Reminder: Your sample '{product}' is scheduled for testing on {test_date}.\nPlease prepare in advance."

                    try:
                        yag.send(to=owner_email, subject=subject, contents=body)
                        print(f"Email sent to {owner_email}")
                    except Exception as e:
                        print(f"Failed to send email to {owner_email}: {e}")
        except Exception as e:
            print(f"Error checking file {file}: {e}")

def manage_users():
    clear_tree(tree_ref, tree_frame_ref)
    users = load_users()

    tree_frame_ref[0] = tk.Frame(right_panel, bg="white")
    tree_frame_ref[0].pack(expand=True, fill="both", padx=10, pady=10)

    columns = ["Username", "Password", "Role"]
    tree = ttk.Treeview(tree_frame_ref[0], columns=columns, show="headings")
    tree_ref[0] = tree

    for col in columns:
        tree.heading(col, text=col)
        tree.column(col, width=150, anchor="center")

    def refresh_users():
        tree.delete(*tree.get_children())
        for uname, data in users.items():
            tree.insert("", "end", values=(uname, data["password"], data["role"]))

    def add_user():
        uname = simpledialog.askstring("Add User", "Enter new username:")
        if not uname or uname in users:
            messagebox.showerror("Error", "Username exists or invalid.")
            return
        pwd = simpledialog.askstring("Password", "Enter password:", show='*')
        role = simpledialog.askstring("Role", "Enter role (admin/user):")
        if role not in ["admin", "user"]:
            messagebox.showerror("Error", "Role must be 'admin' or 'user'.")
            return
        users[uname] = {"password": pwd, "role": role}
        save_users(users)
        refresh_users()

    def edit_user():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select User", "Please select a user to edit.")
            return
        item = tree.item(sel[0])
        uname = item["values"][0]
        if uname == "admin":
            messagebox.showwarning("Not Allowed", "Cannot edit default admin.")
            return
        new_pwd = simpledialog.askstring("Edit Password", f"New password for {uname}:", show='*')
        new_role = simpledialog.askstring("Edit Role", "New role (admin/user):")
        if new_pwd:
            users[uname]["password"] = new_pwd
        if new_role in ["admin", "user"]:
            users[uname]["role"] = new_role
        save_users(users)
        refresh_users()

    def del_user():
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Select User", "Please select a user to delete.")
            return
        uname = tree.item(sel[0])["values"][0]
        if uname == "admin":
            messagebox.showwarning("Not Allowed", "Cannot delete default admin.")
            return
        confirm = messagebox.askyesno("Confirm Delete", f"Delete user '{uname}'?")
        if confirm:
            del users[uname]
            save_users(users)
            refresh_users()

    # Buttons
    btn_frame = tk.Frame(tree_frame_ref[0])
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="Add User", command=add_user).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Edit User", command=edit_user).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Delete User", command=del_user).pack(side="left", padx=5)

    refresh_users()
    tree.pack(expand=True, fill="both")



# GUI Setup
root = tk.Tk()
icon = tk.PhotoImage(file="transparent.png")
root.iconphoto(False, icon)
root.title(f"Shelf-life System - {role.title()}")
root.geometry("1000x650")

left_panel = tk.Frame(root, bg="lightgray", width=150)
left_panel.pack(side="left", fill="y")

right_panel = tk.Frame(root, bg="white")
right_panel.pack(side="right", expand=True, fill="both")

style = ttk.Style()
style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)

# Buttons (left panel)
ttk.Button(left_panel, text="Load File", style="Bold.TButton", command=load_excel_file).pack(pady=(10, 3), padx=8, fill="x")
ttk.Button(left_panel, text="View All", style="Bold.TButton", command=view_all).pack(pady=3, padx=8, fill="x")
if not is_admin:
    ttk.Button(left_panel, text="My Requests", style="Bold.TButton", command=view_my_requests).pack(pady=3, padx=8, fill="x")
if not is_admin:
    ttk.Button(left_panel, text="Request Barcode", style="Bold.TButton", command=request_barcode_from_selection).pack(pady=3, padx=8, fill="x")

if is_admin:
    ttk.Button(left_panel, text="Approve Row Requests", style="Bold.TButton",
               command=lambda: view_requests(SCRIPT_DIR, right_panel, tree_ref, tree_frame_ref, clear_tree)).pack(pady=3, padx=8, fill="x")
    ttk.Button(left_panel, text="Barcode Requests", style="Bold.TButton",
               command=lambda: generate_barcode_requests(SCRIPT_DIR, right_panel, tree_ref, tree_frame_ref,clear_tree)).pack(pady=3, padx=8, fill="x")
    ttk.Button(left_panel, text="Send Email Reminders", style="Bold.TButton", command=send_email_reminders).pack(pady=3, padx=8, fill="x")
    ttk.Button(left_panel, text="Manage Users", style="Bold.TButton", command=manage_users).pack(pady=3, padx=8, fill="x")
    ttk.Button(left_panel, text="Scan Barcode", style="Bold.TButton", command=test_open_barcode).pack(pady=3, padx=8, fill="x")


# Top action buttons
top_bar = tk.Frame(right_panel, bg="white")
top_bar.pack(fill="x", padx=8, pady=5)

ttk.Button(top_bar, text="Add", style="Bold.TButton", width=15, command=add_data_to_selected_file).pack(side="right", padx=5)
ttk.Button(top_bar, text="Delete", style="Bold.TButton", width=15, command=delete_selected_data).pack(side="right", padx=5)

view_all()
root.mainloop()
