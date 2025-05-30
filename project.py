import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd

def button_clicked(name):
    content_label.config(text=f"{name} button clicked!")

# Create main window
root = tk.Tk()
icon = tk.PhotoImage(file="transparent.png")
root.iconphoto(False, icon)
root.title("Test")
root.geometry("1000x650")

# Create frames for left and right panels
left_panel = tk.Frame(root, bg="lightgray", width=150)
left_panel.pack(side="left", fill="y")

right_panel = tk.Frame(root, bg="white")
right_panel.pack(side="right", expand=True, fill="both")

# Add buttons to left panel
btn1 = ttk.Button(left_panel, text="Button 1", width=20, command=lambda: button_clicked("Button 1"))
btn1.pack(pady=(10,3), padx=8, fill="x")

btn2 = ttk.Button(left_panel, text="Button 2", command=lambda: button_clicked("Button 2"))
btn2.pack(pady=3, padx=8, fill="x")

btn3 = ttk.Button(left_panel, text="Button 3", command=lambda: button_clicked("Button 3"))
btn3.pack(pady=3, padx=8, fill="x")

# Add a label in the right panel for content
content_label = tk.Label(right_panel, text="Welcome!", font=("Arial", 16), bg="white")
content_label.pack(pady=20)

# Start the main event loop
root.mainloop()
