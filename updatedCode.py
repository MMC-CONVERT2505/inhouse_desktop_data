import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import pandas as pd
import os

# ---- Loading Modal ----
loading_modal = None

def show_loading(message="Please wait..."):
    global loading_modal
    if loading_modal is not None:
        return  # already open

    loading_modal = tk.Toplevel(root)
    loading_modal.title("Loading")
    loading_modal.geometry("250x100")
    loading_modal.resizable(False, False)
    loading_modal.transient(root)  # stay on top
    loading_modal.grab_set()       # make modal

    ttk.Label(loading_modal, text=message, anchor="center").pack(pady=10)
    pb = ttk.Progressbar(loading_modal, mode="indeterminate")
    pb.pack(fill="x", padx=20, pady=10)
    pb.start(10)

    # Disable close button
    loading_modal.protocol("WM_DELETE_WINDOW", lambda: None)

    # Center the modal
    root.update_idletasks()
    x = root.winfo_x() + (root.winfo_width() // 2) - (250 // 2)
    y = root.winfo_y() + (root.winfo_height() // 2) - (100 // 2)
    loading_modal.geometry(f"+{x}+{y}")

def hide_loading():
    global loading_modal
    if loading_modal:
        loading_modal.destroy()
        loading_modal = None


DEFAULT_DSN = "QuickBooks Data"
conn = None
company_name = "Unknown"

def connect(dsn_name: str):
    global conn, company_name
    show_loading("Connecting...") 
    status_var.set("Connecting...")
    root.update_idletasks()

    conn = pyodbc.connect(f"DSN={dsn_name};", autocommit=True)
    try:
        df = pd.read_sql("SELECT CompanyName FROM Company", conn)
        if not df.empty:
            company_name = df.iloc[0, 0].replace(" ", "_")
    except Exception:
        company_name = "Unknown"
    status_var.set(f"Connected: {company_name}")
    hide_loading()
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
# def fetch_coa(connection):
#     try:
#         return pd.read_sql("SELECT Name, AccountNumber, AccountType, Balance, IsActive FROM Account", connection)
#     except Exception:
#         return pd.read_sql("SELECT * FROM Account", connection)
def fetch_coa(connection):
    try:
        query = """SELECT * FROM Account"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Account", connection)

# def fetch_customers(connection):
#     try:
#         query = """SELECT ListID, Name, CompanyName, Phone, Email, IsActive FROM Customer"""
#         return pd.read_sql(query, connection)
#     except Exception:
#         return pd.read_sql("SELECT * FROM Customer", connection)

# def fetch_suppliers(connection):
#     try:
#         query = """SELECT ListID, Name, CompanyName, Phone, Email, IsActive FROM Vendor"""
#         return pd.read_sql(query, connection)
#     except Exception:
#         return pd.read_sql("SELECT * FROM Vendor", connection)

# def fetch_items(connection):
#     try:
#         query = """SELECT ListID, Name, FullName, SalesDesc, SalesPrice, IsActive FROM Item"""
#         return pd.read_sql(query, connection)
#     except Exception:
#         return pd.read_sql("SELECT * FROM Item", connection)

# def fetch_journal(connection):
#     try:
#         query = """SELECT TxnID, TxnDate, RefNumber, AccountRefFullName, Memo, Amount FROM JournalEntry"""
#         return pd.read_sql(query, connection)
#     except Exception:
#         return pd.read_sql("SELECT * FROM JournalEntry", connection)

def fetch_bill(connection):
    try:
        query = """SELECT * FROM Bill"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Bill", connection)

def fetch_invoice(connection):
    try:
        query = """SELECT * FROM Invoice"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM Invoice", connection)
    
def fetch_receivepayment(connection):
    try:
        query = """SELECT * FROM ReceivePaymentLine"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM ReceivePaymentLine", connection)

def fetch_BillPaymentCheckLine(connection):
    try:
        query = """SELECT * FROM BillPaymentCheckLine"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM BillPaymentCheckLine", connection)
    
def fetch_BillPaymentCreditCardLine(connection):
    try:
        query = """SELECT * FROM BillPaymentCreditCardLine"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM BillPaymentCreditCardLine", connection)
    
def fetch_InvoiceLine(connection):
    try:
        query = """SELECT * FROM InvoiceLine"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM InvoiceLine", connection)
    
def fetch_CreditMemoLinkedTxn(connection):
    try:
        query = """SELECT * FROM CreditMemoLinkedTxn"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM CreditMemoLinkedTxn", connection)
    
def fetch_VendorCreditLinkedTxn(connection):
    try:
        query = """SELECT * FROM VendorCreditLinkedTxn"""
        return pd.read_sql(query, connection)
    except Exception:
        return pd.read_sql("SELECT * FROM VendorCreditLinkedTxn", connection)            
       
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
        show_loading(f"Exporting {data_type}...")   
        with connect(dsn) as connection:
            status_var.set(f"Exporting {data_type}...")
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

        status_var.set("Export complete.")
        hide_loading()
        messagebox.showinfo("Done", f"Exported {data_type} to:\n{save_path}")
        try:
            os.startfile(save_path)
        except Exception:
            pass
    except Exception as e:
        status_var.set("Export failed.")
        hide_loading()
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
        show_loading(f"Exporting All Data...")   
        with connect(dsn) as connection:
            status_var.set("Exporting all datasets...")
            root.update_idletasks()

            datasets = {
                "ChartOfAccounts": fetch_coa(connection),
                # "Customers": fetch_customers(connection),
                # "Suppliers": fetch_suppliers(connection),
                # "Items": fetch_items(connection),
                "Bill": fetch_bill(connection),
                "Invoice": fetch_invoice(connection),
                "ReceivePayment": fetch_receivepayment(connection),
                "BillPaymentCheckLine": fetch_BillPaymentCheckLine(connection),
                "BillPaymentCreditCardLine": fetch_BillPaymentCreditCardLine(connection),
            }

            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                workbook = writer.book
                for sheet_name, df in datasets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)

                    worksheet = writer.sheets[sheet_name]
                    for idx, col in enumerate(df.columns):
                        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(idx, idx, max_len)

        status_var.set("All exports complete.")
        hide_loading()
        messagebox.showinfo("Done", f"All datasets exported to:\n{save_path}")
        try:
            os.startfile(save_path)
        except Exception:
            pass
    except Exception as e:
        status_var.set("Export failed.")
        hide_loading()
        messagebox.showerror("Export failed", str(e))

# ---- UI ----
root = tk.Tk()
root.title("Reckon Data Exporter")
root.geometry("540x400")

frm = ttk.Frame(root, padding=12)
frm.pack(fill="both", expand=True)

ttk.Label(frm, text="DSN (32-bit):").grid(row=0, column=0, sticky="w")
dsn_var = tk.StringVar(value=DEFAULT_DSN)
dsn_combo = ttk.Combobox(frm, textvariable=dsn_var, width=34)
dsn_combo.grid(row=0, column=1, sticky="we", padx=8)
frm.columnconfigure(1, weight=1)

ttk.Button(frm, text="Test Connection", command=lambda: connect(dsn_var.get())).grid(row=0, column=2, padx=4)

ttk.Separator(frm).grid(row=1, column=0, columnspan=3, pady=10, sticky="ew")

ttk.Label(frm, text="Export:").grid(row=2, column=0, sticky="w")
ttk.Button(frm, text="Export All", command=export_all).grid(row=2, column=1, sticky="w", pady=10)
ttk.Button(frm, text="Chart of Accounts", command=lambda: export_data(fetch_coa, "ChartOfAccounts")).grid(row=3, column=1, sticky="w", pady=6)
# ttk.Button(frm, text="Customers", command=lambda: export_data(fetch_customers, "Customers")).grid(row=3, column=1, sticky="w", pady=6)
# ttk.Button(frm, text="Suppliers", command=lambda: export_data(fetch_suppliers, "Suppliers")).grid(row=4, column=1, sticky="w", pady=6)
# ttk.Button(frm, text="Items", command=lambda: export_data(fetch_items, "Items")).grid(row=4, column=2, sticky="w", pady=6)
ttk.Button(frm, text="Bill", command=lambda: export_data(fetch_bill, "Bill")).grid(row=3, column=2, sticky="w", pady=6)
ttk.Button(frm, text="Invoice", command=lambda: export_data(fetch_invoice, "Invoice")).grid(row=4, column=1, sticky="w", pady=6)
ttk.Button(frm, text="Receive Payment Line", command=lambda: export_data(fetch_receivepayment, "ReceivePaymentLine")).grid(row=4, column=2, sticky="w", pady=6)
ttk.Button(frm, text="BillPayment Check Line", command=lambda: export_data(fetch_BillPaymentCheckLine, "BillPaymentCheckLine")).grid(row=5, column=1, sticky="w", pady=6)
ttk.Button(frm, text="BillPayment Credit Card Line", command=lambda: export_data(fetch_BillPaymentCreditCardLine, "BillPaymentCreditCardLine")).grid(row=5, column=2, sticky="w", pady=6)

ttk.Separator(frm).grid(row=6, column=0, columnspan=3, pady=10, sticky="ew")

ttk.Button(frm, text="Invoice Line", command=lambda: export_data(fetch_InvoiceLine, "InvoiceLine")).grid(row=7, column=1, sticky="w", pady=6)
ttk.Button(frm, text="Credit Memo Linked Txn", command=lambda: export_data(fetch_CreditMemoLinkedTxn, "CreditMemoLinkedTxn")).grid(row=7, column=2, sticky="w", pady=6)
ttk.Button(frm, text="Vendor Credit Linked Txn", command=lambda: export_data(fetch_VendorCreditLinkedTxn, "VendorCreditLinkedTxn")).grid(row=8, column=1, sticky="w", pady=6)

ttk.Separator(frm).grid(row=9, column=0, columnspan=3, pady=10, sticky="ew")

status_var = tk.StringVar(value="Idle")
ttk.Label(frm, textvariable=status_var, foreground="blue").grid(row=10, column=0, columnspan=5, pady=(12,0), sticky="w")

def on_close():
    close_connection()
    root.destroy()

root.protocol("WM_DELETE_WINDOW", on_close)
root.mainloop()
