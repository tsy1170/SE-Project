import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import messagebox
import bcrypt
import admin_page

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-beb7183669.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

# ----------------- GUI Setup -----------------
def login():
    admin_id = entry_admin_id.get()
    password = entry_password.get()

    if not admin_id or not password:
        messagebox.showwarning("Input Error", "Please enter both UserID and Password.")
        return

    # Retrieve document by UserID
    admin_ref = db.collection("Admin").document(admin_id)
    admin_doc = admin_ref.get()

    if admin_doc.exists:
        admin_data = admin_doc.to_dict()
        stored_password = admin_data.get("Password")

        # Compare password
        if password == stored_password:
            # messagebox.showinfo("Login Success", f"Welcome, {user_data.get('Username')}!")
            root.destroy()
            admin_page.admin_panel(admin_data)
        else:
            messagebox.showerror("Login Failed", "Incorrect password.")
    else:
        messagebox.showerror("Login Failed", "UserID not found.")


# Tkinter GUI
root = tk.Tk()
root.title("Admin Login")
root.geometry("300x200")

tk.Label(root, text="AdminID:").pack(pady=5)
entry_admin_id = tk.Entry(root)
entry_admin_id.pack(pady=5)

tk.Label(root, text="Password:").pack(pady=5)
entry_password = tk.Entry(root, show="*")
entry_password.pack(pady=5)

tk.Button(root, text="Login", command=login).pack(pady=15)

root.mainloop()
