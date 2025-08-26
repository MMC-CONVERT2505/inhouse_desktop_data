import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import pandas as pd
import os

DEFAULT_DSN = "QuickBooks Data"
conn = None
company_name = "Unknown"

def connect(dsn_name: str):
    global conn, company_name
    status_var.set("Connecting...")
    progress.start(10)
    root.update_idletasks()

    conn = pyodbc.connect(f"DSN={dsn_name};", autocommit=True)
    try:
        df = pd.read_sql("SELECT CompanyName FROM Company", conn)
        if not df.empty:
            company_name = df.iloc[0, 0].replace(" ", "_")
    except Exception:
        company_name = "Unknown"

    progress.stop()
    status_var.set(f"Connected: {company_name}")
    return conn

def close_connection():
    global conn
    if conn:
        try:
            conn.close()
        except:
            pass
        conn = None
        status_var.set("Connection closed.")

# ---- Data fetch functions ----
def fetch_coa(connection):
    try:
        return pd.read_sql("SELECT Name, AccountNumber, AccountType, Balance, IsActive FROM Account", connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Account", connection)

def fetch_customers(connection):
    try:
        return pd.read_sql("SELECT ListID, Name, CompanyName, Phone, Email, IsActive FROM Customer", connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Customer", connection)

def fetch_suppliers(connection):
    try:
        return pd.read_sql("SELECT ListID, Name, CompanyName, Phone, Email, IsActive FROM Vendor", connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Vendor", connection)

def fetch_items(connection):
    try:
        return pd.read_sql("SELECT ListID, Name, FullName, SalesDesc, SalesPrice, IsActive FROM Item", connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Item", connection)

# ---- Data fetch for Journal / All Transactions ----
def fetch_journal(connection):
    try:
        # Try JournalEntry (actual QB table)
        return pd.read_sql(
            "SELECT TxnID, TxnDate, RefNumber, AccountRefFullName, Memo, Amount FROM JournalEntry",
            connection
        )
    except Exception:
        # Fallback: dump all transaction data
        return pd.read_sql("SELECT * FROM Transaction", connection)

   
# ---- Export single dataset ----
def export_data(fetch_func, data_type):
    dsn = dsn_var.get().strip()
    if not dsn:
        messagebox.showerror("Error", "Please enter an ODBC DSN name.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile=f"{company_name}_{data_type}.xlsx",
        title="Save export as"
    )
    if not save_path:
        return

    try:
        with connect(dsn) as connection:
            status_var.set(f"Exporting {data_type}...")
            progress.start(10)
            root.update_idletasks()

            df = fetch_func(connection)
            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name=data_type, index=False)

                # formatting
                workbook = writer.book
                worksheet = writer.sheets[data_type]
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)

        progress.stop()
        status_var.set("Export complete.")
        messagebox.showinfo("Done", f"Exported {data_type} to:\n{save_path}")
        try:
            os.startfile(save_path)
        except Exception:
            pass
    except Exception as e:
        progress.stop()
        status_var.set("Export failed.")
        messagebox.showerror("Export failed", str(e))

# ---- Export all datasets in one file ----
def export_all():
    dsn = dsn_var.get().strip()
    if not dsn:
        messagebox.showerror("Error", "Please enter an ODBC DSN name.")
        return

    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile=f"{company_name}_AllData.xlsx",
        title="Save export as"
    )
    if not save_path:
        return

    try:
        with connect(dsn) as connection:
            status_var.set("Exporting all datasets...")
            progress.start(10)
            root.update_idletasks()

            datasets = {
                "ChartOfAccounts": fetch_coa(connection),
                "Customers": fetch_customers(connection),
                "Suppliers": fetch_suppliers(connection),
                "Items": fetch_items(connection),
            }

            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                workbook = writer.book
                for sheet_name, df in datasets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    worksheet = writer.sheets[sheet_name]
                    for idx, col in enumerate(df.columns):
                        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(idx, idx, max_len)

        progress.stop()
        status_var.set("All exports complete.")
        messagebox.showinfo("Done", f"All datasets exported to:\n{save_path}")
        try:
            os.startfile(save_path)
        except Exception:
            pass
    except Exception as e:
        progress.stop()
        status_var.set("Export failed.")
        messagebox.showerror("Export failed", str(e))

# ---- UI ----
root = tk.Tk()
root.title("Reckon (QODBC) Exporter")
root.geometry("540x400")

frm = ttk.Frame(root, padding=12)
frm.pack(fill="both", expand=True)

ttk.Label(frm, text="ODBC DSN (32-bit):").grid(row=0, column=0, sticky="w")
dsn_var = tk.StringVar(value=DEFAULT_DSN)
dsn_combo = ttk.Combobox(frm, textvariable=dsn_var, width=34)
dsn_combo.grid(row=0, column=1, sticky="we", padx=8)
frm.columnconfigure(1, weight=1)

ttk.Button(frm, text="Test Connection", command=lambda: connect(dsn_var.get())).grid(row=0, column=2, padx=4)

ttk.Separator(frm).grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")

ttk.Label(frm, text="Export:").grid(row=2, column=0, sticky="w")
ttk.Button(frm, text="Chart of Accounts", command=lambda: export_data(fetch_coa, "ChartOfAccounts")).grid(row=2, column=1, sticky="w", pady=6)
ttk.Button(frm, text="Customers", command=lambda: export_data(fetch_customers, "Customers")).grid(row=3, column=1, sticky="w", pady=6)
ttk.Button(frm, text="Suppliers", command=lambda: export_data(fetch_suppliers, "Suppliers")).grid(row=4, column=1, sticky="w", pady=6)
ttk.Button(frm, text="Items", command=lambda: export_data(fetch_items, "Items")).grid(row=5, column=1, sticky="w", pady=6)
ttk.Button(frm, text="All Transactions (Journal)", command=lambda: export_data(fetch_journal, "Journal")).grid(row=6, column=1, sticky="w", pady=6)

ttk.Button(frm, text="Export All", command=export_all).grid(row=7, column=1, sticky="w", pady=10)

status_var = tk.StringVar(value="Idle")
ttk.Label(frm, textvariable=status_var, foreground="blue").grid(row=8, column=0, columnspan=3, pady=(12,0), sticky="w")

progress = ttk.Progressbar(frm, mode="indeterminate")
progress.grid(row=9, column=0, columnspan=3, pady=(8,0), sticky="we")

def on_close():
    close_connection()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
