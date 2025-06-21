import tkinter as tk
from tkinter import messagebox
import firebase_admin
from firebase_admin import credentials, firestore

import admin_page
import user_page

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-12d8d15ff2.json")  # your Firebase JSON
if not firebase_admin._apps:
    firebase_admin.initialize_app(cred)
db = firestore.client()

# ----------------- Login Logic -----------------
def login(root, entry_ID, entry_password):
    ID = entry_ID.get().strip()
    password = entry_password.get().strip()

    if not ID or not password:
        messagebox.showwarning("Missing Info", "Please enter both ID and Password.")
        return

    # Check Admin
    admin_doc = db.collection("Admin").document(ID).get()
    if admin_doc.exists:
        admin_data = admin_doc.to_dict()
        if admin_data["Password"] == password:
            messagebox.showinfo("Login Success", f"Welcome Admin: {ID}")
            root.destroy()
            admin_page.admin_panel(admin_data, db)
            return
        else:
            messagebox.showerror("Error", "Incorrect admin password.")
            return

    # Check User
    user_doc = db.collection("Users").document(ID).get()
    if user_doc.exists:
        user_data = user_doc.to_dict()
        if user_data["Password"] == password:
            messagebox.showinfo("Login Success", f"Welcome User: {ID}")
            root.destroy()
            user_page.user_panel(user_data, db)
            return
        else:
            messagebox.showerror("Error", "Incorrect user password.")
            return

    # If not found
    messagebox.showerror("Login Failed", "Account not found.")


# ----------------- Login GUI -----------------
def show_login():
    root = tk.Tk()
    root.title("Login")
    root.geometry("300x200")

    tk.Label(root, text="ID:").pack(pady=5)
    entry_ID = tk.Entry(root)
    entry_ID.pack(pady=5)

    tk.Label(root, text="Password:").pack(pady=5)
    entry_password = tk.Entry(root, show="*")
    entry_password.pack(pady=5)

    tk.Button(root, text="Login", width=10,
              command=lambda: login(root, entry_ID, entry_password)).pack(pady=20)

    root.mainloop()

# Run on file start
if __name__ == "__main__":
    show_login()
