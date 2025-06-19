import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import messagebox
import user_page

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-beb7183669.json")
firebase_admin.initialize_app(cred)
db = firestore.client()

# ----------------- GUI Setup -----------------
def login():
    user_id = entry_userID.get()
    password = entry_password.get()

    if not user_id or not password:
        messagebox.showwarning("Input Error", "Please enter both UserID and Password.")
        return

    # Retrieve document by UserID
    user_ref = db.collection("Users").document(user_id)
    user_doc = user_ref.get()

    if user_doc.exists:
        user_data = user_doc.to_dict()
        stored_password = user_data.get("Password")

        # Compare password
        if password == stored_password:
            # messagebox.showinfo("Login Success", f"Welcome, {user_data.get('Username')}!")
            root.destroy()
            user_page.user_panel(user_data, db)
        else:
            messagebox.showerror("Login Failed", "Incorrect password.")
    else:
        messagebox.showerror("Login Failed", "UserID not found.")


# Tkinter GUI
root = tk.Tk()
root.title("User Login")
root.geometry("300x200")

tk.Label(root, text="UserID:").pack(pady=5)
entry_userID = tk.Entry(root)
entry_userID.pack(pady=5)

tk.Label(root, text="Password:").pack(pady=5)
entry_password = tk.Entry(root, show="*")
entry_password.pack(pady=5)

tk.Button(root, text="Login", command=login).pack(pady=15)

root.mainloop()
