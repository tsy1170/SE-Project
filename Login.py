import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import messagebox
import user_page
import admin_page

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-beb7183669.json")
if not firebase_admin._apps:
    firebase_admin.initialize_app(cred)
db = firestore.client()


def login(root, entry_ID, entry_password):
    ID = entry_ID.get()
    password = entry_password.get()

    if not ID or not password:
        messagebox.showwarning("Input Error", "Please enter both UserID and Password.")
        return

    # Check Users collection first
    user_ref = db.collection("Users").document(ID)
    user_doc = user_ref.get()

    if user_doc.exists:
        user_data = user_doc.to_dict()
        if password == user_data.get("Password"):
            root.destroy()
            user_page.user_panel(user_data, db)
        else:
            messagebox.showerror("Login Failed", "Incorrect password.")
        return

    # Check Admin collection
    admin_ref = db.collection("Admin").document(ID)
    admin_doc = admin_ref.get()

    if admin_doc.exists:
        admin_data = admin_doc.to_dict()
        if password == admin_data.get("Password"):
            root.destroy()
            admin_page.admin_panel(admin_data, db)
        else:
            messagebox.showerror("Login Failed", "Incorrect password.")
        return

    messagebox.showerror("Login Failed", "UserID not found in Users or Admin collections.")

def show_login():
    root = tk.Tk()
    root.title("Login")
    root.geometry("300x200")

    tk.Label(root, text="UserID:").pack(pady=5)
    entry_ID = tk.Entry(root)
    entry_ID.pack(pady=5)

    tk.Label(root, text="Password:").pack(pady=5)
    entry_password = tk.Entry(root, show="*")
    entry_password.pack(pady=5)

    tk.Button(root, text="Login", command=lambda: login(root, entry_ID, entry_password)).pack(pady=15)

    root.mainloop()


if __name__ == "__main__":
    show_login()
