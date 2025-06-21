import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import Login

tree = None
tree_frame = None
top_bar = None


def logout(root):
    global top_bar
    top_bar = None
    root.destroy()
    Login.show_login()

def tester_panel(tester_data, db):
    root = tk.Tk()
    root.title(f"Welcome, {tester_data.get("TesterName")}")
    root.geometry("1000x650")

    # Create frames for left and right panels
    left_panel = tk.Frame(root, bg="lightgray", width=150)
    left_panel.pack(side="left", fill="y")

    right_panel = tk.Frame(root, bg="white")
    right_panel.pack(side="right", expand=True, fill="both")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)
    style.configure("Treeview.Heading", background="#d3d3d3", foreground="black", font=("Segoe UI", 10, "bold"))

    # Add buttons to left panel
    btn_load_file = ttk.Button(left_panel, text="Load File", style="Bold.TButton")
    btn_load_file.pack(pady=(10,3), padx=8, fill="x")

    ttk.Button(left_panel, text="Logout", style="Bold.TButton", command=lambda: logout(root)).pack(pady=20, padx=10, side="bottom")

    # Start the main event loop
    root.mainloop()
