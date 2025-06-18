# --- Admin Functions Module (admin_function.py) ---
import os
import json
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from barcode import Code128
from barcode.writer import ImageWriter

def get_requests_file(script_dir):
    return os.path.join(script_dir, "requests.json")

def init_requests(script_dir):
    requests_file = get_requests_file(script_dir)
    if not os.path.exists(requests_file):
        with open(requests_file, "w") as f:
            json.dump({"requests": []}, f)

def load_requests(script_dir):
    with open(get_requests_file(script_dir), "r") as f:
        return json.load(f)

def save_requests(script_dir, data):
    with open(get_requests_file(script_dir), "w") as f:
        json.dump(data, f, indent=2)

def view_requests(script_dir, right_panel, tree_ref, tree_frame_ref, clear_tree):
    clear_tree(tree_ref, tree_frame_ref)
    requests = load_requests(script_dir)
    data = [r for r in requests["requests"] if r.get("status", "pending") == "pending"]

    if not data:
        messagebox.showinfo("No Requests", "No pending row-level requests found.")
        return

    tree_frame_ref[0] = tk.Frame(right_panel, bg="white")
    tree_frame_ref[0].pack(expand=True, fill="both", padx=10, pady=10)

    columns = ["username", "action", "file", "status", "timestamp", "data"]
    tree_ref[0] = ttk.Treeview(tree_frame_ref[0], columns=columns, show="headings")

    for col in columns:
        tree_ref[0].heading(col, text=col)
        tree_ref[0].column(col, width=120, anchor="center")

    for i, req in enumerate(data):
        tree_ref[0].insert("", "end", iid=i, values=(
            req.get("username", ""),
            req.get("action", ""),
            req.get("file", ""),
            req.get("status", "pending"),
            req.get("timestamp", ""),
            str(req.get("data", {}))
        ))

    tree_ref[0].pack(expand=True, fill="both")

    def approve():
        sel = tree_ref[0].selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a request.")
            return

        for item in sel:
            idx = int(item)
            request = data[idx]
            real_index = requests["requests"].index(request)
            try:
                file_path = os.path.join(script_dir, request["file"])
                df = pd.read_excel(file_path, engine="openpyxl")
                if request["action"] == "add":
                    df.loc[len(df)] = request["data"]
                elif request["action"] == "delete":
                    df = df[~(df.astype(str) == pd.Series(request["data"], index=df.columns).astype(str)).all(axis=1)]
                df.to_excel(file_path, index=False, engine="openpyxl")
                requests["requests"][real_index]["status"] = "approved"
            except Exception as e:
                messagebox.showerror("Error", f"Failed to apply request:\n{e}")

        save_requests(script_dir, requests)
        view_requests(script_dir, right_panel, tree_ref, tree_frame_ref, clear_tree)

    def reject():
        sel = tree_ref[0].selection()
        if not sel:
            messagebox.showwarning("No Selection", "Please select a request.")
            return

        for item in sel:
            idx = int(item)
            request = data[idx]
            real_index = requests["requests"].index(request)
            requests["requests"][real_index]["status"] = "rejected"

        save_requests(script_dir, requests)
        view_requests(script_dir, right_panel, tree_ref, tree_frame_ref, clear_tree)

    btn_frame = tk.Frame(tree_frame_ref[0], bg="white")
    btn_frame.pack(pady=5)
    ttk.Button(btn_frame, text="Approve", command=approve).pack(side="left", padx=5)
    ttk.Button(btn_frame, text="Reject", command=reject).pack(side="left", padx=5)


def generate_barcode_requests(script_dir, right_panel, tree_ref, tree_frame_ref, clear_tree):
    clear_tree(tree_ref, tree_frame_ref)
    init_requests(script_dir)
    all_data = load_requests(script_dir)
    barcode_requests = [(i, r) for i, r in enumerate(all_data["requests"]) if r["action"] == "barcode" and r["status"] == "pending"]

    if not barcode_requests:
        messagebox.showinfo("No Barcodes", "No pending barcode requests.")
        return

    tree_frame_ref[0] = tk.Frame(right_panel, bg="white")
    tree_frame_ref[0].pack(expand=True, fill="both", padx=10, pady=10)

    columns = ["index", "username", "file", "data"]
    tree_ref[0] = ttk.Treeview(tree_frame_ref[0], columns=columns, show="headings")

    for col in columns:
        tree_ref[0].heading(col, text=col)
        tree_ref[0].column(col, width=150, anchor="center")

    for idx, req in barcode_requests:
        tree_ref[0].insert("", "end", iid=str(idx), values=(idx, req["username"], req["file"], str(req["data"])))

    tree_ref[0].pack(expand=True, fill="both")

    def approve():
        barcode_dir = os.path.join(script_dir, "barcodes")
        os.makedirs(barcode_dir, exist_ok=True)

        for idx, req in barcode_requests:
            try:
                barcode_content = f"{req['file']}::{json.dumps(req['data'])}"
                barcode_img = Code128(barcode_content, writer=ImageWriter())
                file_name = f"barcode_{idx}.png"
                full_path = os.path.join(barcode_dir, file_name)
                barcode_img.save(full_path)
                all_data["requests"][idx]["status"] = "approved"
            except Exception as e:
                messagebox.showerror("Error", f"Failed to generate barcode for request {idx}:\n{e}")

        save_requests(script_dir, all_data)
        messagebox.showinfo("Done", "Barcodes generated and saved.")
        generate_barcode_requests(script_dir, right_panel, tree_ref, tree_frame_ref, clear_tree)

    ttk.Button(tree_frame_ref[0], text="Approve All", command=approve).pack(pady=10)
