import tkinter as tk
from tkinter import messagebox
import firebase_admin
from firebase_admin import credentials, firestore

import admin_page
import user_page
import tester_page

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-fcfd17746c.json")  # your Firebase JSON
if not firebase_admin._apps:
    firebase_admin.initialize_app(cred)
db = firestore.client()

# ----------------- Login Logic -----------------
def login(root, entry_ID, entry_password):
    ID = entry_ID.get()
    password = entry_password.get()

    if not ID or not password:
        messagebox.showerror("Error", "Please enter both ID and password.")
        return

    try:
        # First, check Admin
        admin_doc = db.collection("Admin").document(ID).get()
        if admin_doc.exists and admin_doc.to_dict()["Password"] == password:
            root.destroy()
            admin_page.admin_panel(admin_doc.to_dict(), db)
            return

        # Then check Users
        user_doc = db.collection("Users").document(ID).get()
        if user_doc.exists and user_doc.to_dict()["Password"] == password:
            root.destroy()
            user_page.user_panel(user_doc.to_dict(), db)
            return

        # Then check Tester
        tester_doc = db.collection("Tester").document(ID).get()
        if tester_doc.exists and tester_doc.to_dict()["Password"] == password:
            root.destroy()
            tester_page.tester_panel(tester_doc.to_dict(), db)
            return

        # If not found in any
        messagebox.showerror("Login Failed", "Invalid ID or password.")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")


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
