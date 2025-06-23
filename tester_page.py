import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from tkcalendar import DateEntry
import pandas as pd
import logging
import Login


tree = None
tree_frame = None
top_bar = None


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


def update_cell_in_excel_and_firestore(db, excel_file_path, product_id, submitted_at, column_name, new_value):
    try:
        df = pd.read_excel(excel_file_path, engine="openpyxl")
        df.loc[df["Product_ID"].astype(str) == str(product_id), column_name] = new_value
        df.to_excel(excel_file_path, index=False)

        doc_ref = db.collection(submitted_at).document(str(product_id))
        doc_ref.update({column_name: new_value})
        return True, ""
    except Exception as e:
        logging.error(f"Update Error: {e}")
        return False, str(e)


def open_date_picker(db, excel_file_path, column_name, title):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row.")
        return

    item = selected[0]
    row_values = tree.item(item, "values")
    columns = tree["columns"]

    def submit_date():
        new_date = cal.get_date().strftime("%d-%m-%Y")

        updated_values = list(row_values)
        start_idx = columns.index(column_name)
        updated_values[start_idx] = new_date
        tree.item(item, values=updated_values)

        product_id = updated_values[columns.index("Product_ID")]
        submitted_at = updated_values[columns.index("Submitted_At")]

        success, msg = update_cell_in_excel_and_firestore(db, excel_file_path, product_id, submitted_at, column_name, new_date)
        if success:
            messagebox.showinfo("Success", f"{column_name.replace('_', ' ')} updated successfully.")
        else:
            messagebox.showerror("Update Error", msg)

        date_win.destroy()

    date_win = tk.Toplevel()
    date_win.configure(bg="White")
    date_win.title(title)
    date_win.geometry("250x140")
    date_win.grab_set()

    ttk.Label(date_win, text=f"Select {column_name.replace('_', ' ')}:", style="Custom.TLabel").pack(pady=(10, 5))
    cal = DateEntry(date_win, date_pattern="dd-mm-yyyy")
    cal.pack(pady=5)

    ttk.Button(date_win, text="Submit", command=submit_date).pack(pady=10)


def load_file(right_panel, db):
    global tree, tree_frame, top_bar

    file_path = filedialog.askopenfilename(title="Select Excel File", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, engine="openpyxl")

        if "Submitted_At" not in df.columns:
            messagebox.showerror("Missing Column", "The file must include 'Submitted_At' column.")
            return


        clear_right_panel(right_panel)

        if top_bar and top_bar.winfo_exists():
            top_bar.destroy()
        top_bar = None

        top_bar = tk.Frame(right_panel, bg="white")
        top_bar.pack(fill="x", padx=8, pady=5)

        ttk.Button(top_bar, text="Test End", style="Bold.TButton", width=15,
                   command=lambda: open_date_picker(db, file_path, "Test_End_Date", "Select Test End Date")).pack(side="right", pady=(10, 0), padx=5)

        ttk.Button(top_bar, text="Test Start", style="Bold.TButton", width=15,
                   command=lambda: open_date_picker(db, file_path, "Test_Start_Date", "Select Test Start Date")).pack(side="right", pady=(10, 0), padx=5)

        ttk.Label(top_bar, text="🔴 <15 days  🟡 15–30 days", background="white", font=("Segoe UI", 9)).pack(side="left", padx=10)

        search_var = tk.StringVar()
        search_entry = ttk.Entry(top_bar, textvariable=search_var, width=30)
        search_entry.pack(side="left", padx=(10, 0), pady=(10, 0))

        def filter_treeview():
            query = search_var.get().lower()
            for item in tree.get_children():
                values = tree.item(item, "values")
                if any(query in str(val).lower() for val in values):
                    tree.reattach(item, "", "end")
                else:
                    tree.detach(item)

        search_entry.bind("<KeyRelease>", lambda e: filter_treeview())

        tree_frame = tk.Frame(right_panel, bg="white")
        tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

        tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings", selectmode="browse")
        tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_y = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")

        scrollbar_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")

        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, anchor="center")

        tree.delete(*tree.get_children())

        for _, row in df.iterrows():
            row_values = [row[col] if col not in ["Test_Start_Date", "Test_End_Date"] or pd.notna(row[col]) and str(row[col]).strip() != "" else "-" for col in df.columns]
            tag = ""
            try:
                test_end_str = row.get("Test_End_Date", "")
                if isinstance(test_end_str, str) and test_end_str != "-" and test_end_str.strip():
                    test_end_date = datetime.strptime(test_end_str, "%d-%m-%Y").date()
                    days_left = (test_end_date - datetime.today().date()).days
                    if 0 <= days_left <= 15:
                        tag = "critical"
                    elif 15 < days_left <= 30:
                        tag = "warning"
            except Exception as e:
                logging.error(f"Tagging Error: {e}")
            tree.insert("", "end", values=row_values, tags=(tag,))

        tree.tag_configure("warning", background="yellow")
        tree.tag_configure("critical", background="tomato")

        for _, row in df.iterrows():
            try:
                submitted_at = row.get("Submitted_At")
                date_obj = datetime.strptime(str(submitted_at), "%d-%m-%Y")
                collection_name = date_obj.strftime("%d-%m-%Y")
                product_id = str(row.get("Product_ID", "unknown"))
                doc_data = row.to_dict()
                db.collection(collection_name).document(product_id).set(doc_data)
            except Exception as e:
                logging.error(f"Firestore Sync Error: {e}")

        df.to_excel(file_path, index=False)
        messagebox.showinfo("Success", "File loaded.")

    except Exception as e:
        logging.error(f"Load File Error: {e}")
        messagebox.showerror("Error", f"Failed to load file:\n{e}")


def logout(root):
    global top_bar
    top_bar = None
    root.destroy()
    Login.show_login()


def tester_panel(tester_data, db):
    root = tk.Tk()
    root.title(f"Welcome, {tester_data.get('TesterName')}")
    root.geometry("1000x650")

    left_panel = tk.Frame(root, bg="lightgray", width=150)
    left_panel.pack(side="left", fill="y")

    right_panel = tk.Frame(root, bg="white")
    right_panel.pack(side="right", expand=True, fill="both")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Bold.TButton", font=("Segoe UI", 10, "bold"), width=20, border=15)
    style.configure("Treeview.Heading", background="#d3d3d3", foreground="black", font=("Segoe UI", 10, "bold"))
    style.configure("Custom.TLabel", background="White", foreground="#333333", font=("Segoe UI", 10, "bold"), padding=5)

    ttk.Button(left_panel, text="Load File", style="Bold.TButton", command=lambda: load_file(right_panel, db)).pack(pady=(10, 3), padx=8, fill="x")

    ttk.Button(left_panel, text="Logout", style="Bold.TButton", command=lambda: logout(root)).pack(pady=20, padx=10, side="bottom")

    root.mainloop()
