import firebase_admin
from firebase_admin import credentials, firestore
import tkinter as tk
from tkinter import messagebox, ttk
import user_page
import admin_page
import tester_page

# ----------------- Firestore Setup -----------------
cred = credentials.Certificate("se-project-ad0dd-firebase-adminsdk-fbsvc-40f9620543.json")
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

    # Check Tester collection
    tester_ref = db.collection("Tester").document(ID)
    tester_doc = tester_ref.get()

    if tester_doc.exists:
        tester_data = tester_doc.to_dict()
        if password == tester_data.get("Password"):
            root.destroy()
            tester_page.tester_panel(tester_data, db)
        else:
            messagebox.showerror("Login Failed", "Incorrect password.")
        return

    messagebox.showerror("Login Failed", "ID not found.")


def show_login():
    root = tk.Tk()
    root.title("Shelf Life Study Management - Login")
    root.geometry("500x400")
    root.configure(bg="#f0f2f5")
    root.minsize(400, 350)  # allow resizing
    root.resizable(True, True)

    # Styling
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TLabel", background="white", font=("Segoe UI", 10))
    style.configure("Header.TLabel", font=("Segoe UI", 16, "bold"), foreground="#2A4F6E", background="white")
    style.configure("Custom.TEntry", font=("Segoe UI", 10), foreground="black", fieldbackground="lightyellow", padding=5)
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"))

    # Hover and active style map for button
    style.map("Bold.TButton",
              background=[("active", "#3b7dd8"), ("!disabled", "#4a90e2")],
              foreground=[("disabled", "#ccc")])

    # Outer Frame
    card = tk.Frame(root, bg="white", bd=1, relief="solid")
    card.pack(expand=True, fill="both", padx=30, pady=30)

    # Header
    ttk.Label(card, text="Shelf Life Study Management", style="Header.TLabel").pack(pady=(20, 10))

    # Form
    form_frame = tk.Frame(card, bg="white")
    form_frame.pack(pady=10, padx=20)

    ttk.Label(form_frame, text="User ID:", background="white").grid(row=0, column=0, sticky="w", pady=5)
    entry_ID = ttk.Entry(form_frame, style="Custom.TEntry", width=30)
    entry_ID.grid(row=1, column=0, pady=(0, 10))

    ttk.Label(form_frame, text="Password:", background="white").grid(row=2, column=0, sticky="w", pady=5)
    entry_password = ttk.Entry(form_frame, show="*", style="Custom.TEntry", width=30)
    entry_password.grid(row=3, column=0, pady=(0, 5))

    # Show Password Checkbox
    show_password_var = tk.BooleanVar()
    def toggle_password():
        entry_password.config(show="" if show_password_var.get() else "*")

    tk.Checkbutton(form_frame, text="Show Password", variable=show_password_var, onvalue=True, offvalue=False,
                   command=toggle_password, bg="white", font=("Segoe UI", 9)).grid(row=4, column=0, sticky="w", pady=5)

    # Login Button
    login_btn = ttk.Button(card, text="Login", style="Bold.TButton",
                           command=lambda: login(root, entry_ID, entry_password))
    login_btn.pack(pady=20)

    # Footer
    footer = tk.Label(root, text="© 2025 SLMS | Version 1.0", bg="#f0f2f5", fg="#888", font=("Segoe UI", 8))
    footer.pack(side="bottom", pady=5)

    root.mainloop()

if __name__ == "__main__":
    show_login()
