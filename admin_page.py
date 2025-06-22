import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import re
import os
import platform
import subprocess
import smtplib
import ssl
from email.message import EmailMessage
from datetime import datetime, timedelta
import barcode
from barcode import get as get_barcode
from barcode.writer import ImageWriter
import Login


tree = None
tree_frame = None
top_bar = None


def edit_user(right_panel, db):
    global tree

    selected = tree.selection()
    if not selected or len(selected) != 1:
        messagebox.showwarning("Selection Error", "Please select exactly one user to edit.")
        return

    item = selected[0]
    values = tree.item(item, "values")
    user_id, username, email, password = values

    form = tk.Toplevel(right_panel)
    form.configure(bg="White")
    form.title(f"Edit User: {user_id}")
    form.geometry("350x300")
    form.grab_set()

    labels = ["User ID", "Username", "Email", "Password"]
    entries = {}

    # Create label-entry pairs
    for i, label in enumerate(labels):
        ttk.Label(form, text=label+":", style="Custom.TLabel").grid(row=i, column=0, padx=(23, 10), pady=10, sticky="e")
        entry = ttk.Entry(form, width=30)
        entry.grid(row=i, column=1, padx=10, pady=10)
        entries[label] = entry

    # Fill in existing values
    entries["User ID"].insert(0, user_id)
    entries["User ID"].config(state="disabled")
    entries["Username"].insert(0, username)
    entries["Email"].insert(0, email)
    entries["Password"].insert(0, password)

    def submit_update():
        updated_data = {
            "Username": entries["Username"].get().strip(),
            "Email": entries["Email"].get().strip(),
            "Password": entries["Password"].get().strip()
        }

        if not all(updated_data.values()):
            messagebox.showwarning("Missing Info", "All fields are required.")
            form.lift()
            form.focus_force()
            return

        email_pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        if not re.match(email_pattern, updated_data["Email"]):
            messagebox.showerror("Invalid Email", "Please enter a valid email address.")
            form.lift()
            form.focus_force()
            return

        try:
            db.collection("Users").document(user_id).update(updated_data)
            messagebox.showinfo("Success", "User information updated.")
            form.destroy()
            display_all_users(right_panel, db)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to update user:\n{e}")

    ttk.Button(form, text="Update", command=submit_update).grid(row=5, column=0, columnspan=2, pady=15, padx=(32, 0))


def delete_user(right_panel, db):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a user to delete.")
        return

    confirm = messagebox.askyesno("Confirm Delete", "Are you sure you want to delete the selected user(s)?")
    if not confirm:
        return

    try:
        for item in selected:
            values = tree.item(item, "values")
            user_id = values[0]

            # Delete from Firestore
            db.collection("Users").document(user_id).delete()

            # Remove from Treeview
            tree.delete(item)

        messagebox.showinfo("Deleted", "Selected user(s) have been deleted.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to delete user(s):\n{e}")


def add_user(right_panel, db):
    form = tk.Toplevel(right_panel)
    form.configure(bg="White")
    form.title("Add New User")
    form.geometry("350x300")
    form.grab_set()

    # Labels and Entries
    labels = ["User ID", "Username", "Email", "Password"]
    entries = {}

    for i, label in enumerate(labels):
        ttk.Label(form, text=label+":", style="Custom.TLabel").grid(row=i, column=0, padx=(23, 10), pady=10, sticky="e")
        entry = ttk.Entry(form, width=30)
        entry.grid(row=i, column=1, padx=10, pady=10)
        entries[label] = entry


    def submit_user():
        user_id = entries["User ID"].get().strip()
        username = entries["Username"].get().strip()
        email = entries["Email"].get().strip()
        password = entries["Password"].get().strip()

        if not user_id or not username or not email or not password:
            messagebox.showwarning("Missing Info", "All fields are required.")
            form.lift()
            form.focus_force()
            return

        email_pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        if not re.match(email_pattern, email):
            messagebox.showerror("Invalid Email", "Please enter a valid email address.")
            form.lift()
            form.focus_force()
            return

        try:
            # Write to Firestore
            db.collection("Users").document(user_id).set({
                "Username": username,
                "Email": email,
                "Password": password,
                "Joined": datetime.now()
            })
            messagebox.showinfo("Success", "User added successfully.")
            form.destroy()
            display_all_users(right_panel, db)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to add user:\n{e}")

    # Submit button
    ttk.Button(form, text="Submit", command=submit_user).grid(row=len(labels), column=0, columnspan=2, pady=20, padx=(32, 0))


def display_all_users(right_panel, db):
    global top_bar, tree, tree_frame

    clear_right_panel(right_panel)

    if top_bar is not None:
        try:
            if top_bar.winfo_exists():
                top_bar.destroy()
        except tk.TclError:
            pass
        top_bar = None

    if top_bar is None:
        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)

        btn_add_user = ttk.Button(
            top_bar, text="Add User", style="Green.TButton", width=15,
            command=lambda: add_user(right_panel, db)
        )
        btn_add_user.pack(side="right", pady=(10, 0), padx=5)

        btn_delete_user = ttk.Button(
            top_bar, text="Delete User", style="Red.TButton", width=15,
            command=lambda: delete_user(right_panel, db)
        )
        btn_delete_user.pack(side="right", pady=(10, 0), padx=5)

        btn_edit_user = ttk.Button(
            top_bar, text="Edit User", style="Bold.TButton", width=15,
            command=lambda: edit_user(right_panel, db)
        )
        btn_edit_user.pack(side="right", pady=(10, 0), padx=5)

    # Create Treeview
    if not tree_frame:
        tree_frame = tk.Frame(right_panel, bg="white")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    if not tree:
        tree = ttk.Treeview(tree_frame,
                            columns=("User ID", "Username", "Email", "Password"),
                            show="headings")
        tree.heading("User ID", text="User ID")
        tree.heading("Username", text="Username")
        tree.heading("Email", text="Email")
        tree.heading("Password", text="Password")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # Load users from Firestore
    try:
        users = db.collection("Users").stream()

        for user in users:
            data = user.to_dict()
            tree.insert("", "end", values=(
                user.id,
                data.get("Username", ""),
                data.get("Email", ""),
                data.get("Password", "")
            ))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load users:\n{e}")


def load_barcode(right_panel):
    global tree, tree_frame

    barcode_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Barcode")
    if not os.path.exists(barcode_dir):
        messagebox.showinfo("No Barcode Folder", "The 'Barcode' folder does not exist.")
        return

    files = [f for f in os.listdir(barcode_dir) if f.lower().endswith((".png", ".jpg", ".jpeg"))]
    if not files:
        messagebox.showinfo("No Files", "No barcode images found in the 'Barcode' folder.")
        return

    clear_right_panel(right_panel)

    # tree view
    if not tree_frame:
        tree_frame = tk.Frame(right_panel, bg="white")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

    if not tree:
        tree = ttk.Treeview(tree_frame, columns="File", show="headings")
        tree.heading("File", text="Barcode")
        tree.column("File", anchor="center")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)


    for f in files:
        tree.insert("", "end", values=(f,))

    def open_selected_file(event=None):
        selected = tree.selection()
        if not selected:
            return

        item = tree.item(selected[0])
        file_name = item["values"][0]
        file_path = os.path.join(barcode_dir, file_name)

        try:
            if platform.system() == "Windows":
                os.startfile(file_path)
            elif platform.system() == "Darwin":  # macOS
                subprocess.call(["open", file_path])
            else:  # Linux
                subprocess.call(["xdg-open", file_path])
        except Exception as e:
            messagebox.showerror("Error", f"Could not open file:\n{e}")

    tree.bind("<Double-1>", open_selected_file)


def export_and_send_reminders(db):
    confirm = messagebox.askyesno("Send Reminders", "Are you sure you want to send reminders to users and testers?")
    if not confirm:
        return

    today = datetime.today()
    two_months_after = today + timedelta(days=60)
    date_pattern = re.compile(r"\d{2}-\d{2}-\d{4}")

    user_rows = {}

    # Step 1: Collect matching rows grouped by UserID
    for collection in db.collections():
        col_name = collection.id
        if date_pattern.fullmatch(col_name):
            try:
                docs = db.collection(col_name).stream()
                for doc in docs:
                    data = doc.to_dict()
                    test_date_str = data.get("Test_Date", "")
                    user_id = data.get("UserID", "")

                    try:
                        test_date = datetime.strptime(test_date_str, "%d-%m-%Y")
                        if today < test_date <= two_months_after:
                            row = {
                                "Collection": col_name,
                                "Document_ID": doc.id,
                                **data
                            }
                            if user_id:
                                user_rows.setdefault(user_id, []).append(row)
                    except ValueError:
                        continue
            except Exception as e:
                print(f"Error reading collection {col_name}: {e}")

    # Step 2: Email & Export per user
    for user_id, rows in user_rows.items():
        try:
            user_doc = db.collection("Users").document(user_id).get()
            user_data = user_doc.to_dict()
            if not user_data or "Email" not in user_data:
                print(f"No email found for UserID: {user_id}")
                continue

            user_email = user_data["Email"]

            # Save to Reminders folder
            reminders_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Reminders")
            os.makedirs(reminders_dir, exist_ok=True)

            df = pd.DataFrame(rows)
            df = df[["Product_ID", "Product_Name", "Description", "Test_Date"]]

            file_name = f"Reminder ({user_id}).xlsx"
            file_path = os.path.join(reminders_dir, file_name)
            df.to_excel(file_path, index=False)

            # Send to user
            send_email_with_attachment(
                receiver=user_email,
                subject="Upcoming Product Test Reminder",
                body=f"Hi {user_data.get('Username', '')},\n\nPlease find your product test reminders attached.\n\nBest regards,\nAdmin",
                attachment_path=file_path
            )
            print(f"Sent reminder to {user_email}")

            # Step 3: Also send to all testers
            testers = db.collection("Tester").stream()
            for tester in testers:
                tester_data = tester.to_dict()
                tester_email = tester_data.get("Email")
                if tester_email:
                    send_email_with_attachment(
                        receiver=tester_email,
                        subject=f"Reminder copy for UserID: {user_id}",
                        body=f"Hello Tester,\n\nAttached is the reminder Excel file for UserID {user_id}.\n\nRegards,\nAdmin",
                        attachment_path=file_path
                    )
                    print(f"Also sent to tester: {tester_email}")

        except Exception as e:
            print(f"Failed for UserID {user_id}: {e}")

    messagebox.showinfo("Successful", "Reminders sent.")


def send_email_with_attachment(receiver, subject, body, attachment_path):
    sender_email = "tsy1170@gmail.com"
    app_password = "khvvzmuhytpinkxe"

    msg = EmailMessage()
    msg["From"] = sender_email
    msg["To"] = receiver
    msg["Subject"] = subject
    msg.set_content(body)

    with open(attachment_path, "rb") as f:
        file_data = f.read()
        file_name = os.path.basename(attachment_path)
        msg.add_attachment(file_data, maintype="application", subtype="octet-stream", filename=file_name)

    context = ssl._create_unverified_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(sender_email, app_password)
        server.send_message(msg)


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
        messagebox.showwarning("No Selection", "Please select row(s) to approve.")
        return

    try:
        all_data = []
        submitted_at = None
        test_date = None
        user_id = None

        for item in selected:
            values = tree.item(item, "values")
            product_id, name, desc, test_date_str, submitted_at_str, user_id = values

            if not submitted_at_str:
                messagebox.showerror("Error", "Missing 'Submitted At' field.")
                return

            # Parse date strings
            submitted_at = datetime.strptime(submitted_at_str, "%d-%m-%Y")
            test_date = datetime.strptime(test_date_str, "%d-%m-%Y")
            collection_name = submitted_at.strftime("%d-%m-%Y")

            # Prepare Firestore data
            data = {
                "Product_ID": product_id,
                "Product_Name": name,
                "Description": desc,
                "Test_Date": test_date.strftime("%d-%m-%Y"),
                "Submitted_At": submitted_at.strftime("%d-%m-%Y"),
                "UserID": user_id
            }

            # Write to Firestore
            db.collection(collection_name).document(product_id).set(data)
            db.collection("Pending").document(product_id).delete()

            # For Excel
            data.pop("UserID")
            all_data.append(data)

            # Remove from Treeview
            tree.delete(item)

        # Save to Excel file
        file_name = f"{collection_name}.xlsx"
        file_path = os.path.join(os.getcwd(), file_name)

        df = pd.DataFrame(all_data)
        if os.path.exists(file_path):
            existing_df = pd.read_excel(file_path)
            df = pd.concat([existing_df, df], ignore_index=True)
        df.to_excel(file_path, index=False)

        # Barcode content and file name
        barcode_content = f"{submitted_at.strftime('%d%m%Y')}-{test_date.strftime('%d%m%Y')}"
        barcode_filename = f"Batch_{submitted_at.strftime('%d%m%Y')}_{user_id}"

        barcode_dir = os.path.join(os.getcwd(), "Barcode")
        os.makedirs(barcode_dir, exist_ok=True)
        barcode_path = os.path.join(barcode_dir, barcode_filename)

        code128 = get_barcode("code128", barcode_content, writer=ImageWriter())
        saved_path = code128.save(barcode_path)

        open_barcode_file(saved_path)
        messagebox.showinfo("Success", "Selected request(s) approved")

    except Exception as e:
        messagebox.showerror("Error", f"Approval failed:\n{e}")


def open_barcode_file(saved_path):
    if platform.system() == "Windows":
        os.startfile(saved_path)
    elif platform.system() == "Darwin":  # macOS
        subprocess.run(["open", saved_path])
    else:  # Linux
        subprocess.run(["xdg-open", saved_path])


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


def create_pending_tree_view(right_panel):
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

        # ✅ Green Approve button
        btn_approve = ttk.Button(top_bar, text="Approve", style="Green.TButton", width=15, command=lambda: approve_requests(db))
        btn_approve.pack(side="right", pady=(10, 0), padx=5)

        # ❌ Red Reject button
        btn_reject = ttk.Button(top_bar, text="Reject", style="Red.TButton", width=15, command=lambda: reject_requests(db))
        btn_reject.pack(side="right", pady=(10, 0), padx=5)

    create_pending_tree_view(right_panel)

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
                data.get("Submitted_At", "").strftime("%d-%m-%Y") if isinstance(data.get("Submitted_At"), datetime) else "",
                data.get("UserID", "")
            ))
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load pending data:\n{e}")

def logout(root):
    global top_bar
    top_bar = None
    root.destroy()
    Login.show_login()

def admin_panel(admin_data, db):
    def set_active_button(active_btn):
        nonlocal active_button
        for btn in sidebar_buttons:
            btn.configure(style="Bold.TButton")
        active_btn.configure(style="Active.TButton")
        active_button = active_btn

    root = tk.Tk()
    root.title(f"Welcome, {admin_data.get('AdminID')}")
    root.geometry("1000x650")

    left_panel = tk.Frame(root, bg="lightgray", width=180)
    left_panel.pack(side="left", fill="y")

    right_panel = tk.Frame(root, bg="white")
    right_panel.pack(side="right", expand=True, fill="both")

    # Styling
    style = ttk.Style()
    style.theme_use("clam")

    # Base button
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)
    style.configure("Treeview.Heading", background="#d3d3d3", foreground="black", font=("Segoe UI", 10, "bold"))
    style.configure("Custom.TLabel", background="White", foreground="#333333", font=("Segoe UI", 10, "bold"), padding=5)

    # Active left panel button
    style.configure("Active.TButton", font=("Segoe UI", 10, "bold"), background="#4CAF50", foreground="white")
    style.map("Active.TButton", background=[("active", "#45a049")])

    # Red delete button
    style.configure("Red.TButton", background="#f44336", foreground="white")
    style.map("Red.TButton", background=[("active", "#d32f2f")])

    # Green add button
    style.configure("Green.TButton", background="#4CAF50", foreground="white")
    style.map("Green.TButton", background=[("active", "#45a049")])

    # Sidebar buttons with tracking
    sidebar_buttons = []

    btn_view = ttk.Button(left_panel, text="View Requests", style="Bold.TButton",
                          command=lambda: [set_active_button(btn_view), view_pending_requests(right_panel, db)])
    btn_view.pack(pady=5, padx=10)
    sidebar_buttons.append(btn_view)

    btn_reminders = ttk.Button(left_panel, text="Send Reminders", style="Bold.TButton",
                               command=lambda: [set_active_button(btn_reminders), export_and_send_reminders(db)])
    btn_reminders.pack(pady=5, padx=10)
    sidebar_buttons.append(btn_reminders)

    btn_barcode = ttk.Button(left_panel, text="All Barcode", style="Bold.TButton",
                             command=lambda: [set_active_button(btn_barcode), load_barcode(right_panel)])
    btn_barcode.pack(pady=5, padx=10)
    sidebar_buttons.append(btn_barcode)

    btn_manage = ttk.Button(left_panel, text="Manage Users", style="Bold.TButton",
                            command=lambda: [set_active_button(btn_manage), display_all_users(right_panel, db)])
    btn_manage.pack(pady=5, padx=10)
    sidebar_buttons.append(btn_manage)

    # Logout button
    ttk.Button(left_panel, text="Logout", style="Bold.TButton", command=lambda: logout(root)).pack(
        pady=20, padx=10, side="bottom")

    # Track active button
    active_button = None

    root.mainloop()
