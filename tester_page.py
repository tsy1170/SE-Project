import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from tkcalendar import DateEntry
import pandas as pd
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


def enter_test_end_date(db, excel_file_path):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row.")
        return

    item = selected[0]
    row_values = tree.item(item, "values")
    columns = tree["columns"]

    def submit_date():
        test_end_date = cal.get_date().strftime("%d-%m-%Y")

        # Update Treeview
        updated_values = list(row_values)
        start_idx = columns.index("Test_End_Date")
        updated_values[start_idx] = test_end_date
        tree.item(item, values=updated_values)

        # Update Excel
        try:
            df = pd.read_excel(excel_file_path, engine="openpyxl")
            product_id = updated_values[columns.index("Product_ID")]
            df.loc[df["Product_ID"].astype(str) == str(product_id), "Test_End_Date"] = test_end_date
            df.to_excel(excel_file_path, index=False)
        except Exception as e:
            messagebox.showerror("Excel Error", f"Failed to update Excel:\n{e}")
            date_win.destroy()
            return

        # Update Firestore
        try:
            submitted_at = updated_values[columns.index("Submitted_At")]
            doc_ref = db.collection(submitted_at).document(str(product_id))
            doc_ref.update({"Test_End_Date": test_end_date})
        except Exception as e:
            messagebox.showerror("Firestore Error", f"Failed to update Firestore:\n{e}")
        else:
            messagebox.showinfo("Success", "Test Start Date updated successfully.")
        date_win.destroy()

    # Popup date picker window
    date_win = tk.Toplevel()
    date_win.configure(bg="White")
    date_win.title("Select Test End Date")
    date_win.geometry("250x120")
    date_win.grab_set()

    ttk.Label(date_win, text="Select Test End Date:").pack(pady=(10, 5))
    cal = DateEntry(date_win, date_pattern="dd-mm-yyyy")
    cal.pack(pady=5)

    ttk.Button(date_win, text="Submit", command=submit_date).pack(pady=10)


def enter_test_start_date(db, excel_file_path):
    global tree

    selected = tree.selection()
    if not selected:
        messagebox.showwarning("No Selection", "Please select a row.")
        return

    item = selected[0]
    row_values = tree.item(item, "values")
    columns = tree["columns"]

    def submit_date():
        test_start_date = cal.get_date().strftime("%d-%m-%Y")

        # Update Treeview
        updated_values = list(row_values)
        start_idx = columns.index("Test_Start_Date")
        updated_values[start_idx] = test_start_date
        tree.item(item, values=updated_values)

        # Update Excel
        try:
            df = pd.read_excel(excel_file_path, engine="openpyxl")
            product_id = updated_values[columns.index("Product_ID")]
            df.loc[df["Product_ID"].astype(str) == str(product_id), "Test_Start_Date"] = test_start_date
            df.to_excel(excel_file_path, index=False)
        except Exception as e:
            messagebox.showerror("Excel Error", f"Failed to update Excel:\n{e}")
            date_win.destroy()
            return

        # Update Firestore
        try:
            submitted_at = updated_values[columns.index("Submitted_At")]
            doc_ref = db.collection(submitted_at).document(str(product_id))
            doc_ref.update({"Test_Start_Date": test_start_date})
        except Exception as e:
            messagebox.showerror("Firestore Error", f"Failed to update Firestore:\n{e}")
        else:
            messagebox.showinfo("Success", "Test Start Date updated successfully.")
        date_win.destroy()

    # Popup date picker window
    date_win = tk.Toplevel()
    date_win.configure(bg="White")
    date_win.title("Select Test Start Date")
    date_win.geometry("250x120")
    date_win.grab_set()

    ttk.Label(date_win, text="Select Test Start Date:").pack(pady=(10, 5))
    cal = DateEntry(date_win, date_pattern="dd-mm-yyyy")
    cal.pack(pady=5)

    ttk.Button(date_win, text="Submit", command=submit_date).pack(pady=10)


def load_file(right_panel, db):
    global tree, tree_frame, top_bar

    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )

    if not file_path:
        return

    try:
        df = pd.read_excel(file_path, engine="openpyxl")

        if "Submitted_At" not in df.columns:
            messagebox.showerror("Missing Column", "The file must include 'Submitted_At' column.")
            return

        # Add new columns if not present
        if "Test_Start_Date" not in df.columns:
            df["Test_Start_Date"] = "-"
        if "Test_End_Date" not in df.columns:
            df["Test_End_Date"] = "-"

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

            btn_test_end = ttk.Button(top_bar, text="Test End", style="Bold.TButton", width=15, command=lambda: enter_test_end_date(db, file_path))
            btn_test_end.pack(side="right", pady=(10, 0), padx=5)

            btn_test_start = ttk.Button(top_bar, text="Test Start", style="Bold.TButton", width=15, command=lambda: enter_test_start_date(db, file_path))
            btn_test_start.pack(side="right", pady=(10, 0), padx=5)

        if not tree_frame:
            tree_frame = tk.Frame(right_panel, bg="white")
            tree_frame.pack(expand=True, fill="both", padx=10, pady=10)

        if not tree:
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

        for _, row in df.iterrows():
            row_values = []
            for col in df.columns:
                value = row[col]
                if col in ["Test_Start_Date", "Test_End_Date"]:
                    row_values.append(value if pd.notna(value) and str(value).strip() != "" else "-")
                else:
                    row_values.append(value)
            tree.insert("", "end", values=row_values)


        # Write each row to Firestore
        for _, row in df.iterrows():
            submitted_at = row.get("Submitted_At")
            try:
                date_obj = datetime.strptime(str(submitted_at), "%d-%m-%Y")
                collection_name = date_obj.strftime("%d-%m-%Y")
                product_id = str(row.get("Product_ID", "unknown"))
                doc_data = row.to_dict()
                db.collection(collection_name).document(product_id).set(doc_data)
            except Exception as e:
                print(f"Error: {e}")

        # Save updated Excel file
        df.to_excel(file_path, index=False)

        messagebox.showinfo("Success", "File loaded.")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to load file:\n{e}")



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

    btn_load_file = ttk.Button(left_panel, text="Load File", style="Bold.TButton", command=lambda: load_file(right_panel, db))
    btn_load_file.pack(pady=(10,3), padx=8, fill="x")

    ttk.Button(left_panel, text="Logout", style="Bold.TButton", command=lambda: logout(root)).pack(pady=20, padx=10, side="bottom")

    root.mainloop()
