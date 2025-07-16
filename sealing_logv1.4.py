import tkinter as tk
from tkinter import ttk
import pandas as pd
import csv
import math
from datetime import datetime

# === Spreadsheet Setup ===
xlsx_path = r"C:\Users\DITTJA\OneDrive - Pearson PLC\Desktop\coding and continuing improve\Sealing Log\seal_pull.xlsx"

tables_df = pd.read_excel(xlsx_path, sheet_name="Tables", engine="openpyxl")

table_values = tables_df["Table"].dropna().astype(str).tolist()

operators_df = pd.read_excel(xlsx_path, sheet_name="Operators")
operator_values = operators_df["Operator"].dropna().astype(str).tolist()

inventory_df = pd.read_excel(xlsx_path, sheet_name="Inventory")
sealed_values = inventory_df["Sealed Inventory"].dropna().astype(str).tolist()

# === GUI Setup ===
root = tk.Tk()
root.title("Sealing App")
root.geometry("800x580")

# === Inventory Section ===
tk.Label(root, text="SEALED INVENTORY ITEM").grid(row=0, column=0, padx=5, pady=5)
sealed_combo = ttk.Combobox(root, values=sealed_values, state="readonly")
sealed_combo.grid(row=0, column=1, padx=5, pady=5)

raw_var = tk.StringVar()
books_var = tk.StringVar()
seals_var = tk.StringVar()
rate_var = tk.StringVar()

tk.Label(root, text="RAW INVENTORY ITEM").grid(row=0, column=2, padx=5, pady=5)
raw_entry = tk.Entry(root, textvariable=raw_var)
raw_entry.grid(row=0, column=3, padx=5, pady=5)

tk.Label(root, text="BOOKS IN STACK").grid(row=1, column=0, padx=5, pady=5)
stack_entry = tk.Entry(root, textvariable=books_var)
stack_entry.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="NUMBER OF SEALS").grid(row=1, column=2, padx=5, pady=5)
seals_entry = tk.Entry(root, textvariable=seals_var)
seals_entry.grid(row=1, column=3, padx=5, pady=5)

tk.Label(root, text="RATE").grid(row=1, column=4, padx=5, pady=5)
rate_entry = tk.Entry(root, textvariable=rate_var)
rate_entry.grid(row=1, column=5, padx=5, pady=5)

# === Auto-fill logic ===
def update_inventory_fields(event):
    selected = sealed_combo.get()
    row = inventory_df[inventory_df["Sealed Inventory"] == selected]
    if not row.empty:
        raw_var.set(row["Raw Inventory"].values[0])
        books_var.set(row["Books per Stack"].values[0])
        seals_var.set(row["Seals per Book"].values[0])
        rate_var.set(row["Expected Rate"].values[0])

sealed_combo.bind("<<ComboboxSelected>>", update_inventory_fields)

# === Operator Row Labels ===
tk.Label(root, text="OPERATOR").grid(row=2, column=0)
tk.Label(root, text="TABLE").grid(row=2, column=1)
tk.Label(root, text="LOG IN").grid(row=2, column=2)
tk.Label(root, text="LUNCH").grid(row=2, column=3)
tk.Label(root, text="STACKS").grid(row=2, column=4)
tk.Label(root, text="BOOKS").grid(row=2, column=5)
tk.Label(root, text="LOG OUT").grid(row=2, column=6)
tk.Label(root, text="COMMENTS").grid(row=2, column=7)



# === Prepare entry lists and time tracking ===
operator_entries = []
operator_boxes = []
table_boxes = []
login_times = [""] * 5
logout_times = [""] * 5
lunch_checks = []
comment_boxes = []

def log_in_action(operator_id):
    login_times[operator_id] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Operator {operator_id + 1} logged in at {login_times[operator_id]}")

def log_out_action(operator_id):
    logout_times[operator_id] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"Operator {operator_id + 1} logged out at {logout_times[operator_id]}")

def submit_data():
    sealed_item  = sealed_combo.get()
    raw_item     = raw_var.get()
    books_stack  = books_var.get()
    seals_book   = seals_var.get()
    rate         = rate_var.get()

    rows = []
    for i in range(len(operator_entries)):
        op_name      = operator_boxes[i].get()
        table_name   = table_boxes[i].get()
        stacks_value = operator_entries[i][2].get()
        books_value  = operator_entries[i][3].get()
        comments = comment_boxes[i].get()

        if not any([op_name, table_name, stacks_value, books_value]):
            continue

        login_time_str  = login_times[i]
        logout_time_str = logout_times[i]

        # ⏱️ Duration calculation
        duration_float = ""
        if login_time_str and logout_time_str:
            login_dt  = datetime.strptime(login_time_str, "%Y-%m-%d %H:%M:%S")
            logout_dt = datetime.strptime(logout_time_str, "%Y-%m-%d %H:%M:%S")
            duration_sec = (logout_dt - login_dt).total_seconds()
            duration_hours = duration_sec / 3600

            # Subtract 0.5 hours if lunch was taken
            if lunch_checks[i].get():
                duration_hours -= 0.5

            
            # Ensure minimum duration of 0.1 hours
            duration_hours = max(duration_hours, 0.1)
            duration_float = round(math.ceil(duration_hours * 10) / 10, 1)



        # 📚 Total Books calculation
        try:
            stacks_int = int(stacks_value)
            books_stack_int = int(books_stack)
            books_int = int(books_value)
            total_books = (stacks_int * books_stack_int) + books_int
        except ValueError:
            total_books = ""

        # 📈 Actual Rate calculation
        actual_rate = ""
        if total_books != "" and duration_float not in ["", 0]:
            actual_rate = round(total_books / duration_float, 2)

        row = [
            login_time_str,
            logout_time_str,
            duration_float,
            op_name,
            table_name,
            stacks_value,
            books_value,
            total_books,
            actual_rate,
            sealed_item,
            raw_item,
            books_stack,
            seals_book,
            rate,
            comments
        ]
        rows.append(row)

    filename = f"production_log_{datetime.now():%Y%m%d_%H%M%S}.csv"
    with open(filename, mode='w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow([
            "Login Time", "Logout Time", "Total Time (hrs)", "Operator", "Table",
            "Stacks", "Books", "Total Books", "Actual Rate",
            "Sealed Inventory", "Raw Inventory", "Books per Stack",
            "Seals per Book", "Expected Rate", "Comments"
        ])
        writer.writerows(rows)

    print(f"Data exported successfully to {filename}")

# === Build Operator Rows ===
for i in range(5):
    row_idx = 3 + i

    op_combo = ttk.Combobox(root, values=operator_values, state="readonly")
    op_combo.grid(row=row_idx, column=0, padx=5, pady=5)
    operator_boxes.append(op_combo)

    table_combo = ttk.Combobox(root, values=table_values, width=4, state="readonly")
    table_combo.grid(row=row_idx, column=1, padx=5, pady=5)
    table_boxes.append(table_combo)

    login_btn = tk.Button(root, text="LOG IN", command=lambda i=i: log_in_action(i))
    login_btn.grid(row=row_idx, column=2, padx=5, pady=5)

    lunch_var = tk.BooleanVar()
    lunch_check = tk.Checkbutton(root, variable=lunch_var)
    lunch_check.grid(row=row_idx, column=3, padx=5, pady=5)
    lunch_checks.append(lunch_var)


    stacks_entry = tk.Entry(root, width=4)  # or width=3
    stacks_entry.grid(row=row_idx, column=4, padx=5, pady=5)

    books_entry = tk.Entry(root, width=4)  # or width=3
    books_entry.grid(row=row_idx, column=5, padx=5, pady=5)


    logout_btn = tk.Button(root, text="LOG OUT", command=lambda i=i: log_out_action(i))
    logout_btn.grid(row=row_idx, column=6, padx=5, pady=5)

    comment_entry = tk.Entry(root, width=30)
    comment_entry.grid(row=row_idx, column=7, padx=5, pady=5)
    comment_boxes.append(comment_entry)


    operator_entries.append((op_combo, table_combo, stacks_entry, books_entry))

# === Submit Button ===
submit_btn = tk.Button(root, text="SUBMIT", command=submit_data)
submit_btn.grid(row=9, column=0, columnspan=6, pady=20)

# === Launch GUI ===
root.mainloop()