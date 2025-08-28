# import win32com.client
# import xml.etree.ElementTree as ET
# import tkinter as tk
# from tkinter import ttk, messagebox, filedialog
# import time
# import xml.sax.saxutils as saxutils
# import pyodbc
# import pandas as pd
# import os

# # ===============================================================
# # ---------------------- COA Numbering Code ----------------------
# # ===============================================================

# def build_account_query(version):
#     return f"""<?xml version="1.0" encoding="utf-8"?>
# <?qbxml version="{version}"?>
# <QBXML>
#   <QBXMLMsgsRq onError="stopOnError">
#     <AccountQueryRq requestID="1">
#       <ActiveStatus>All</ActiveStatus>
#     </AccountQueryRq>
#   </QBXMLMsgsRq>
# </QBXML>"""

# def build_account_mod(version, listid, edit_seq, name, acct_type, new_number):
#     safe_name = saxutils.escape(name)
#     return f"""<?xml version="1.0" encoding="utf-8"?>
# <?qbxml version="{version}"?>
# <QBXML>
#   <QBXMLMsgsRq onError="stopOnError">
#     <AccountModRq>
#       <AccountMod>
#         <ListID>{listid}</ListID>
#         <EditSequence>{edit_seq}</EditSequence>
#         <Name>{safe_name}</Name>
#         <AccountType>{acct_type}</AccountType>
#         <AccountNumber>{new_number}</AccountNumber>
#       </AccountMod>
#     </AccountModRq>
#   </QBXMLMsgsRq>
# </QBXML>"""

# def open_connection():
#     rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
#     rp.OpenConnection2("", "Reckon COA Tool", 1)
#     ticket = rp.BeginSession("", 2)
#     return rp, ticket

# def close_connection_sdk(rp, ticket):
#     rp.EndSession(ticket)
#     rp.CloseConnection()

# def detect_qbxml_version(rp, ticket):
#     try:
#         versions = rp.QBXMLVersionsForSession(ticket)
#         if isinstance(versions, (list, tuple)):
#             return versions[-1]
#         return str(versions)
#     except:
#         return "6.1"

# def fetch_accounts(rp, ticket, version):
#     request = build_account_query(version)
#     response = rp.ProcessRequest(ticket, request)
#     root = ET.fromstring(response)
#     return root.findall(".//AccountRet")

# def auto_numbering(rp, ticket, version, log_callback):
#     accounts = fetch_accounts(rp, ticket, version)
#     used_numbers = [int(acc.findtext("AccountNumber")) for acc in accounts if acc.findtext("AccountNumber") and acc.findtext("AccountNumber").isdigit()]
#     next_number = max(used_numbers)+1 if used_numbers else 1000

#     updated_count = 0
#     for acc in accounts:
#         listid = acc.findtext("ListID")
#         edit_seq = acc.findtext("EditSequence")
#         name = acc.findtext("Name")
#         acct_type = acc.findtext("AccountType")
#         num = acc.findtext("AccountNumber")

#         if not num and edit_seq:
#             try:
#                 rp.ProcessRequest(ticket, build_account_mod(version, listid, edit_seq, name, acct_type, str(next_number)))
#                 log_callback(f"✔ {name} → Assigned #{next_number}")
#                 next_number += 1
#                 updated_count += 1
#                 time.sleep(0.05)
#             except Exception as e:
#                 log_callback(f"❌ Error updating {name}: {e}")
#     return updated_count

# def handle_fetch(log, progress):
#     try:
#         progress.start()
#         log("Fetching Chart of Accounts...")
#         rp, ticket = open_connection()
#         version = detect_qbxml_version(rp, ticket)
#         accounts = fetch_accounts(rp, ticket, version)
#         close_connection_sdk(rp, ticket)
#         log(f"✅ {len(accounts)} accounts fetched!")
#     except Exception as e:
#         log(f"❌ Error: {e}")
#         messagebox.showerror("Error", str(e))
#     finally:
#         progress.stop()

# def handle_auto_number(log, progress):
#     try:
#         progress.start()
#         log("Running Auto Numbering...")
#         rp, ticket = open_connection()
#         version = detect_qbxml_version(rp, ticket)
#         updated = auto_numbering(rp, ticket, version, log)
#         close_connection_sdk(rp, ticket)
#         log(f"✅ {updated} accounts updated with missing numbers!")
#         messagebox.showinfo("Auto Numbering", f"{updated} accounts updated successfully!")
#     except Exception as e:
#         log(f"❌ Error: {e}")
#         messagebox.showerror("Error", str(e))
#     finally:
#         progress.stop()


# # ===============================================================
# # ---------------------- Data Export Code ------------------------
# # ===============================================================

# loading_modal = None
# DEFAULT_DSN = "QuickBooks Data"
# conn = None
# company_name = "Unknown"

# def show_loading(message="Please wait..."):
#     global loading_modal
#     if loading_modal is not None: return
#     loading_modal = tk.Toplevel(root)
#     loading_modal.title("Loading")
#     loading_modal.geometry("250x100")
#     loading_modal.resizable(False, False)
#     loading_modal.transient(root)
#     loading_modal.grab_set()
#     ttk.Label(loading_modal, text=message, anchor="center").pack(pady=10)
#     pb = ttk.Progressbar(loading_modal, mode="indeterminate")
#     pb.pack(fill="x", padx=20, pady=10)
#     pb.start(10)
#     loading_modal.protocol("WM_DELETE_WINDOW", lambda: None)

# def hide_loading():
#     global loading_modal
#     if loading_modal:
#         loading_modal.destroy()
#         loading_modal = None

# def connect(dsn_name: str):
#     global conn, company_name
#     show_loading("Connecting...")
#     status_var.set("Connecting...")
#     root.update_idletasks()
#     conn = pyodbc.connect(f"DSN={dsn_name};", autocommit=True)
#     try:
#         df = pd.read_sql("SELECT CompanyName FROM Company", conn)
#         if not df.empty: company_name = df.iloc[0, 0].replace(" ", "_")
#     except Exception: company_name = "Unknown"
#     status_var.set(f"Connected: {company_name}")
#     hide_loading()
#     return conn

# def close_connection():
#     global conn
#     if conn:
#         try: conn.close()
#         except: pass
#         conn = None
#         status_var.set("Connection closed.")

# def fetch_coa(connection): return pd.read_sql("SELECT * FROM Account", connection)
# def fetch_bill(connection): return pd.read_sql("SELECT * FROM Bill", connection)
# def fetch_invoice(connection): return pd.read_sql("SELECT * FROM Invoice", connection)
# def fetch_receivepayment(connection): return pd.read_sql("SELECT * FROM ReceivePaymentLine", connection)
# def fetch_BillPaymentCheckLine(connection): return pd.read_sql("SELECT * FROM BillPaymentCheckLine", connection)
# def fetch_BillPaymentCreditCardLine(connection): return pd.read_sql("SELECT * FROM BillPaymentCreditCardLine", connection)
# def fetch_InvoiceLine(connection): return pd.read_sql("SELECT * FROM InvoiceLine", connection)
# def fetch_CreditMemoLinkedTxn(connection): return pd.read_sql("SELECT * FROM CreditMemoLinkedTxn", connection)
# def fetch_VendorCreditLinkedTxn(connection): return pd.read_sql("SELECT * FROM VendorCreditLinkedTxn", connection)

# def export_data(fetch_func, data_type):
#     dsn = dsn_var.get().strip()
#     if not dsn:
#         messagebox.showerror("Error", "Please enter DSN.")
#         return
#     save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
#         filetypes=[("Excel Workbook", "*.xlsx")],
#         initialfile=f"{company_name}_{data_type}.xlsx")
#     if not save_path: return
#     try:
#         show_loading(f"Exporting {data_type}...")
#         with connect(dsn) as connection:
#             df = fetch_func(connection)
#             with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
#                 df.to_excel(writer, sheet_name=data_type, index=False)
#                 worksheet = writer.sheets[data_type]
#                 for idx, col in enumerate(df.columns):
#                     max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
#                     worksheet.set_column(idx, idx, max_len)
#         hide_loading()
#         messagebox.showinfo("Done", f"Exported {data_type} to:\n{save_path}")
#     except Exception as e:
#         hide_loading()
#         messagebox.showerror("Export failed", str(e))

# def export_all():
#     dsn = dsn_var.get().strip()
#     if not dsn:
#         messagebox.showerror("Error", "Please enter DSN.")
#         return
#     save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
#         filetypes=[("Excel Workbook", "*.xlsx")],
#         initialfile=f"{company_name}_AllData.xlsx")
#     if not save_path: return
#     try:
#         show_loading("Exporting All Data...")
#         with connect(dsn) as connection:
#             datasets = {
#                 "ChartOfAccounts": fetch_coa(connection),
#                 "Bill": fetch_bill(connection),
#                 "Invoice": fetch_invoice(connection),
#                 "ReceivePayment": fetch_receivepayment(connection),
#                 "BillPaymentCheckLine": fetch_BillPaymentCheckLine(connection),
#                 "BillPaymentCreditCardLine": fetch_BillPaymentCreditCardLine(connection),
#             }
#             with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
#                 for sheet_name, df in datasets.items():
#                     df.to_excel(writer, sheet_name=sheet_name, index=False)
#                     worksheet = writer.sheets[sheet_name]
#                     for idx, col in enumerate(df.columns):
#                         max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
#                         worksheet.set_column(idx, idx, max_len)
#         hide_loading()
#         messagebox.showinfo("Done", f"All datasets exported to:\n{save_path}")
#     except Exception as e:
#         hide_loading()
#         messagebox.showerror("Export failed", str(e))


# # ===============================================================
# # ---------------------- Combined UI -----------------------------
# # ===============================================================

# root = tk.Tk()
# root.title("Reckon Tools Suite")
# root.geometry("550x350")

# notebook = ttk.Notebook(root)
# notebook.pack(fill="both", expand=True)

# # --- Tab 1: COA Numbering ---
# frame_coa = ttk.Frame(notebook)
# notebook.add(frame_coa, text="Numbering")

# log_box = tk.Text(frame_coa, height=15, width=80, bg="#111", fg="white", font=("Consolas", 10))
# log_box.pack(pady=10)

# def log(msg):
#     log_box.insert(tk.END, msg + "\n")
#     log_box.see(tk.END)


# progress = ttk.Progressbar(frame_coa, mode="indeterminate", length=400)
# progress.pack(pady=5)

# btn_frame = ttk.Frame(frame_coa)
# btn_frame.pack()
# ttk.Button(btn_frame, text="📂 Fetch COA", command=lambda: handle_fetch(log, progress)).grid(row=0, column=0, padx=15)
# ttk.Button(btn_frame, text="🔢 COA Numbering", command=lambda: handle_auto_number(log, progress)).grid(row=0, column=1, padx=15)

# # --- Tab 2: Data Export ---
# frame_export = ttk.Frame(notebook)
# notebook.add(frame_export, text="Data Exporter")

# ttk.Label(frame_export, text="Data Source Name:").grid(row=1, column=0, sticky="w")
# dsn_var = tk.StringVar(value=DEFAULT_DSN)
# dsn_combo = ttk.Combobox(frame_export, textvariable=dsn_var, width=34)
# dsn_combo.grid(row=1, column=1, sticky="we", padx=8)
# ttk.Button(frame_export, text="Test Connection", command=lambda: connect(dsn_var.get())).grid(row=1, column=2, padx=4)

# ttk.Separator(frame_export).grid(row=2, column=0, columnspan=4, pady=10, sticky="ew")

# ttk.Button(frame_export, text="Export All", command=export_all).grid(row=3, column=1, pady=6)
# ttk.Button(frame_export, text="Chart of Accounts", command=lambda: export_data(fetch_coa, "ChartOfAccounts")).grid(row=4, column=0, pady=6)
# ttk.Button(frame_export, text="Bill", command=lambda: export_data(fetch_bill, "Bill")).grid(row=4, column=1, pady=6)
# ttk.Button(frame_export, text="Invoice", command=lambda: export_data(fetch_invoice, "Invoice")).grid(row=4, column=2, pady=6)
# ttk.Button(frame_export, text="Receive Payment", command=lambda: export_data(fetch_receivepayment, "ReceivePaymentLine")).grid(row=5, column=0, pady=6)
# ttk.Button(frame_export, text="BillPayment Check Line", command=lambda: export_data(fetch_BillPaymentCheckLine, "BillPaymentCheckLine")).grid(row=5, column=1, pady=6)
# ttk.Button(frame_export, text="BillPayment Credit Card Line", command=lambda: export_data(fetch_BillPaymentCreditCardLine, "BillPaymentCreditCardLine")).grid(row=5, column=2, pady=6)

# ttk.Separator(frame_export).grid(row=6, column=0, columnspan=4, pady=10, sticky="ew")

# ttk.Button(frame_export, text="Invoice Line", command=lambda: export_data(fetch_InvoiceLine, "InvoiceLine")).grid(row=7, column=0, pady=6)
# ttk.Button(frame_export, text="Credit Memo Linked Txn", command=lambda: export_data(fetch_CreditMemoLinkedTxn, "CreditMemoLinkedTxn")).grid(row=7, column=1, pady=6)
# ttk.Button(frame_export, text="Vendor Credit Linked Txn", command=lambda: export_data(fetch_VendorCreditLinkedTxn, "VendorCreditLinkedTxn")).grid(row=7, column=2, pady=6)

# status_var = tk.StringVar(value="Idle")
# ttk.Label(frame_export, textvariable=status_var, foreground="blue").grid(row=8, column=0, columnspan=3, pady=10, sticky="w")

# def on_close():
#     close_connection()
#     root.destroy()

# root.protocol("WM_DELETE_WINDOW", on_close)
# root.mainloop()


# -*- coding: utf-8 -*-
import sys
import os
import time
import win32com.client
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
import pyodbc
import pandas as pd

from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QTextEdit, QPushButton, QProgressBar, QLabel, QComboBox, QFileDialog, QMessageBox, QDialog
)
import qdarkstyle

# ===============================================================
# --------------------- Tkinter-like SHIMS ----------------------
# (preserve same function calls from your backend)
# ===============================================================

MAIN_WINDOW = None  # will be set after MainWindow init

class messagebox:
    @staticmethod
    def showerror(title, text):
        QMessageBox.critical(MAIN_WINDOW, title, text)

    @staticmethod
    def showinfo(title, text):
        QMessageBox.information(MAIN_WINDOW, title, text)

    @staticmethod
    def showwarning(title, text):
        QMessageBox.warning(MAIN_WINDOW, title, text)

class filedialog:
    @staticmethod
    def asksaveasfilename(defaultextension=".xlsx", filetypes=(("Excel Workbook", "*.xlsx"),), initialfile="output.xlsx"):
        # Build Qt filter from tuples like ("Excel Workbook", "*.xlsx")
        filters = ";;".join([f"{desc} ({pattern})" for desc, pattern in filetypes])
        path, _ = QFileDialog.getSaveFileName(MAIN_WINDOW, "Save As", initialfile, filters)
        return path

# status_var & dsn_var shims to preserve .get() / .set()
class _StatusVar:
    def __init__(self):
        self._value = "Idle"
    def set(self, text):
        self._value = text
        if MAIN_WINDOW and MAIN_WINDOW.status_label:
            MAIN_WINDOW.status_label.setText(text)
        QApplication.processEvents()
    def get(self):
        return self._value

class _DsnVar:
    def __init__(self, default_value="QuickBooks Data"):
        self._value = default_value
    def set(self, text):
        self._value = text
        if MAIN_WINDOW and MAIN_WINDOW.dsn_combo:
            # set current text if present, else add
            idx = MAIN_WINDOW.dsn_combo.findText(text)
            if idx == -1:
                MAIN_WINDOW.dsn_combo.addItem(text)
                idx = MAIN_WINDOW.dsn_combo.findText(text)
            MAIN_WINDOW.dsn_combo.setCurrentIndex(idx)
    def get(self):
        if MAIN_WINDOW and MAIN_WINDOW.dsn_combo:
            return MAIN_WINDOW.dsn_combo.currentText()
        return self._value

# ===============================================================
# ---------------------- Backend (unchanged APIs) ---------------
# ===============================================================

# ---------------------- COA Numbering Code ----------------------

def build_account_query(version):
    return f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="{version}"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <AccountQueryRq requestID="1">
      <ActiveStatus>All</ActiveStatus>
    </AccountQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

def build_account_mod(version, listid, edit_seq, name, acct_type, new_number):
    safe_name = saxutils.escape(name)
    return f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="{version}"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <AccountModRq>
      <AccountMod>
        <ListID>{listid}</ListID>
        <EditSequence>{edit_seq}</EditSequence>
        <Name>{safe_name}</Name>
        <AccountType>{acct_type}</AccountType>
        <AccountNumber>{new_number}</AccountNumber>
      </AccountMod>
    </AccountModRq>
  </QBXMLMsgsRq>
</QBXML>"""

def open_connection():
    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    rp.OpenConnection2("", "Reckon COA Tool", 1)
    ticket = rp.BeginSession("", 2)
    return rp, ticket

def close_connection_sdk(rp, ticket):
    rp.EndSession(ticket)
    rp.CloseConnection()

def detect_qbxml_version(rp, ticket):
    try:
        versions = rp.QBXMLVersionsForSession(ticket)
        if isinstance(versions, (list, tuple)):
            return versions[-1]
        return str(versions)
    except:
        return "6.1"

def fetch_accounts(rp, ticket, version):
    request = build_account_query(version)
    response = rp.ProcessRequest(ticket, request)
    root = ET.fromstring(response)
    return root.findall(".//AccountRet")

def auto_numbering(rp, ticket, version, log_callback):
    accounts = fetch_accounts(rp, ticket, version)
    used_numbers = [
        int(acc.findtext("AccountNumber")) for acc in accounts
        if acc.findtext("AccountNumber") and acc.findtext("AccountNumber").isdigit()
    ]
    next_number = max(used_numbers) + 1 if used_numbers else 1000

    updated_count = 0
    for acc in accounts:
        listid = acc.findtext("ListID")
        edit_seq = acc.findtext("EditSequence")
        name = acc.findtext("Name")
        acct_type = acc.findtext("AccountType")
        num = acc.findtext("AccountNumber")

        if not num and edit_seq:
            try:
                rp.ProcessRequest(
                    ticket,
                    build_account_mod(version, listid, edit_seq, name, acct_type, str(next_number))
                )
                log_callback(f"✔ {name} → Assigned #{next_number}")
                next_number += 1
                updated_count += 1
                time.sleep(0.05)
            except Exception as e:
                log_callback(f"❌ Error updating {name}: {e}")
    return updated_count

def handle_fetch(log, progress):
    try:
        # progress.start() equivalent:
        if isinstance(progress, QProgressBar):
            progress.setRange(0, 0)  # busy
        log("Fetching Chart of Accounts...")
        rp, ticket = open_connection()
        version = detect_qbxml_version(rp, ticket)
        accounts = fetch_accounts(rp, ticket, version)
        close_connection_sdk(rp, ticket)
        log(f"✅ {len(accounts)} accounts fetched!")
    except Exception as e:
        log(f"❌ Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        if isinstance(progress, QProgressBar):
            progress.setRange(0, 1)  # done

def handle_auto_number(log, progress):
    try:
        if isinstance(progress, QProgressBar):
            progress.setRange(0, 0)
        log("Running Auto Numbering...")
        rp, ticket = open_connection()
        version = detect_qbxml_version(rp, ticket)
        updated = auto_numbering(rp, ticket, version, log)
        close_connection_sdk(rp, ticket)
        log(f"✅ {updated} accounts updated with missing numbers!")
        messagebox.showinfo("Auto Numbering", f"{updated} accounts updated successfully!")
    except Exception as e:
        log(f"❌ Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        if isinstance(progress, QProgressBar):
            progress.setRange(0, 1)

# ---------------------- Data Export Code ------------------------

loading_modal = None
DEFAULT_DSN = "QuickBooks Data"
conn = None
company_name = "Unknown"

def show_loading(message="Please wait..."):
    global loading_modal
    if loading_modal is not None:
        return
    loading_modal = QDialog(MAIN_WINDOW)
    loading_modal.setWindowTitle("Loading")
    loading_modal.setModal(True)
    layout = QVBoxLayout()
    layout.addWidget(QLabel(message))
    pb = QProgressBar()
    pb.setRange(0, 0)  # indeterminate
    layout.addWidget(pb)
    loading_modal.setLayout(layout)
    loading_modal.setFixedSize(300, 120)
    loading_modal.setWindowFlags(loading_modal.windowFlags() & ~Qt.WindowContextHelpButtonHint)
    loading_modal.show()
    QApplication.processEvents()

def hide_loading():
    global loading_modal
    if loading_modal:
        loading_modal.close()
        loading_modal = None
        QApplication.processEvents()

def connect(dsn_name: str):
    global conn, company_name
    show_loading("Connecting...")
    status_var.set("Connecting...")
    QApplication.processEvents()
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

def fetch_coa(connection): return pd.read_sql("SELECT * FROM Account", connection)
def fetch_bill(connection): return pd.read_sql("SELECT * FROM Bill", connection)
def fetch_invoice(connection): return pd.read_sql("SELECT * FROM Invoice", connection)
def fetch_receivepayment(connection): return pd.read_sql("SELECT * FROM ReceivePaymentLine", connection)
def fetch_BillPaymentCheckLine(connection): return pd.read_sql("SELECT * FROM BillPaymentCheckLine", connection)
def fetch_BillPaymentCreditCardLine(connection): return pd.read_sql("SELECT * FROM BillPaymentCreditCardLine", connection)
def fetch_InvoiceLine(connection): return pd.read_sql("SELECT * FROM InvoiceLine", connection)
def fetch_CreditMemoLinkedTxn(connection): return pd.read_sql("SELECT * FROM CreditMemoLinkedTxn", connection)
def fetch_VendorCreditLinkedTxn(connection): return pd.read_sql("SELECT * FROM VendorCreditLinkedTxn", connection)

def export_data(fetch_func, data_type):
    dsn = dsn_var.get().strip()
    if not dsn:
        messagebox.showerror("Error", "Please enter DSN.")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile=f"{company_name}_{data_type}.xlsx"
    )
    if not save_path:
        return
    try:
        show_loading(f"Exporting {data_type}...")
        with connect(dsn) as connection:
            df = fetch_func(connection)
            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                df.to_excel(writer, sheet_name=data_type, index=False)
                worksheet = writer.sheets[data_type]
                for idx, col in enumerate(df.columns):
                    max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                    worksheet.set_column(idx, idx, max_len)
        hide_loading()
        messagebox.showinfo("Done", f"Exported {data_type} to:\n{save_path}")
    except Exception as e:
        hide_loading()
        messagebox.showerror("Export failed", str(e))

def export_all():
    dsn = dsn_var.get().strip()
    if not dsn:
        messagebox.showerror("Error", "Please enter DSN.")
        return
    save_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Workbook", "*.xlsx")],
        initialfile=f"{company_name}_AllData.xlsx"
    )
    if not save_path:
        return
    try:
        show_loading("Exporting All Data...")
        with connect(dsn) as connection:
            datasets = {
                "ChartOfAccounts": fetch_coa(connection),
                "Bill": fetch_bill(connection),
                "Invoice": fetch_invoice(connection),
                "ReceivePayment": fetch_receivepayment(connection),
                "BillPaymentCheckLine": fetch_BillPaymentCheckLine(connection),
                "BillPaymentCreditCardLine": fetch_BillPaymentCreditCardLine(connection),
                "InvoiceLine": fetch_InvoiceLine(connection),
                "CreditMemoLinkedTxn": fetch_CreditMemoLinkedTxn(connection),
                "VendorCreditLinkedTxn": fetch_VendorCreditLinkedTxn(connection),
            }
            with pd.ExcelWriter(save_path, engine="xlsxwriter") as writer:
                for sheet_name, df in datasets.items():
                    df.to_excel(writer, sheet_name=sheet_name, index=False)
                    worksheet = writer.sheets[sheet_name]
                    for idx, col in enumerate(df.columns):
                        max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                        worksheet.set_column(idx, idx, max_len)
        hide_loading()
        messagebox.showinfo("Done", f"All datasets exported to:\n{save_path}")
    except Exception as e:
        hide_loading()
        messagebox.showerror("Export failed", str(e))

# ===============================================================
# ---------------------- PyQt5 UI -------------------------------
# ===============================================================

class MainWindow(QMainWindow):
    def __init__(self):
        # Global button style (Dark mode + big size)
        self.button_style = """
            QPushButton {
                background-color: #444;
                color: white;
                border-radius: 8px;
                padding: 15px 24px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #666;
            }
        """
        
        super().__init__()
        self.setWindowTitle("MMC Convert | Reckon, QBD")
        self.resize(600, 500)

        # Tabs
        self.tabs = QTabWidget()
        self.setCentralWidget(self.tabs)

        # # Tab 1: Numbering
        # self.tab_numbering = QWidget()
        # self.tabs.addTab(self.tab_numbering, "Numbering")
        # self.init_numbering_tab()

        # # Tab 2: Data Export
        # self.tab_export = QWidget()
        # self.tabs.addTab(self.tab_export, "Data Exporter")
        # self.init_export_tab()

        # Tab 1: Numbering
        self.tab_numbering = QWidget()
        self.tabs.addTab(self.tab_numbering, "Number")
        self.init_numbering_tab()
        
        # Tab 2: Data Export
        self.tab_export = QWidget()
        self.tabs.addTab(self.tab_export, "Data Exporter")
        self.init_export_tab()
        
        # Apply Dark Modern Style to Tabs
        self.tabs.setStyleSheet("""
            QTabWidget::pane { 
                border: 1px solid #444; 
                background: #222; 
                border-radius: 8px; 
                padding: 2px; 
            } 
            QTabBar::tab { 
                background: #333; 
                color: white; 
                padding: 10px 20px; 
                min-height: 25px; 
                border-top-left-radius: 8px; 
                border-top-right-radius: 8px; 
                margin-right: 20px;
                margin-left: 20px; 
                font-size: 11px; 
            } 
            QTabBar::tab:selected { 
                background: #555;
            } 
            QTabBar::tab:hover { 
                background: #444; 
            }
        """)

        # Bind globals
        global MAIN_WINDOW
        MAIN_WINDOW = self

        
    # ----- Numbering Tab -----
    def init_numbering_tab(self):
        layout = QVBoxLayout()

        self.log_box = QTextEdit()
        self.log_box.setReadOnly(True)
        self.log_box.setStyleSheet("background:#111; color:#fff; font-family:Consolas;")
        layout.addWidget(self.log_box)

        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        row = QHBoxLayout()
        btn_fetch = QPushButton("📂 Fetch COA")
        btn_fetch.setStyleSheet(self.button_style)
        btn_fetch.clicked.connect(lambda: handle_fetch(self.log, self.progress))
        row.addWidget(btn_fetch)

        btn_num = QPushButton("🔢 COA Numbering")
        btn_num.setStyleSheet(self.button_style)
        btn_num.clicked.connect(lambda: handle_auto_number(self.log, self.progress))
        row.addWidget(btn_num)

        row.addStretch(1)
        layout.addLayout(row)

        self.tab_numbering.setLayout(layout)

    def log(self, msg):
        self.log_box.append(msg)
        self.log_box.verticalScrollBar().setValue(self.log_box.verticalScrollBar().maximum())

    # ----- Export Tab -----
    def init_export_tab(self):
        grid = QGridLayout()

        grid.addWidget(QLabel("Data Source Name:"), 0, 0)
        self.dsn_combo = QComboBox()
        self.dsn_combo.addItem(DEFAULT_DSN)
        grid.addWidget(self.dsn_combo, 0, 1)

        btn_test = QPushButton("Test Connection")
        btn_test.setStyleSheet(self.button_style)
        btn_test.clicked.connect(lambda: connect(self.dsn_combo.currentText()))
        grid.addWidget(btn_test, 0, 2)

        # Export buttons (preserve same function calls)
        btn_all = QPushButton("Export All")
        btn_all.setStyleSheet(self.button_style)
        btn_all.clicked.connect(export_all)
        grid.addWidget(btn_all, 1, 1)

        btn_coa = QPushButton("Chart of Accounts")
        btn_coa.setStyleSheet(self.button_style)
        btn_coa.clicked.connect(lambda: export_data(fetch_coa, "ChartOfAccounts"))
        grid.addWidget(btn_coa, 2, 0)

        btn_bill = QPushButton("Bill")
        btn_bill.setStyleSheet(self.button_style)
        btn_bill.clicked.connect(lambda: export_data(fetch_bill, "Bill"))
        grid.addWidget(btn_bill, 2, 1)

        btn_invoice = QPushButton("Invoice")
        btn_invoice.setStyleSheet(self.button_style)
        btn_invoice.clicked.connect(lambda: export_data(fetch_invoice, "Invoice"))
        grid.addWidget(btn_invoice, 2, 2)

        btn_recv = QPushButton("Receive Payment")
        btn_recv.setStyleSheet(self.button_style)
        btn_recv.clicked.connect(lambda: export_data(fetch_receivepayment, "ReceivePaymentLine"))
        grid.addWidget(btn_recv, 3, 0)

        btn_bpc = QPushButton("BillPayment Check Line")
        btn_bpc.setStyleSheet(self.button_style)
        btn_bpc.clicked.connect(lambda: export_data(fetch_BillPaymentCheckLine, "BillPaymentCheckLine"))
        grid.addWidget(btn_bpc, 3, 1)

        btn_bpcc = QPushButton("BillPayment Credit Card Line")
        btn_bpcc.setStyleSheet(self.button_style)
        btn_bpcc.clicked.connect(lambda: export_data(fetch_BillPaymentCreditCardLine, "BillPaymentCreditCardLine"))
        grid.addWidget(btn_bpcc, 3, 2)

        btn_invline = QPushButton("Invoice Line")
        btn_invline.setStyleSheet(self.button_style)
        btn_invline.clicked.connect(lambda: export_data(fetch_InvoiceLine, "InvoiceLine"))
        grid.addWidget(btn_invline, 5, 0)

        btn_cmlink = QPushButton("Credit Memo Linked Txn")
        btn_cmlink.setStyleSheet(self.button_style)
        btn_cmlink.clicked.connect(lambda: export_data(fetch_CreditMemoLinkedTxn, "CreditMemoLinkedTxn"))
        grid.addWidget(btn_cmlink, 5, 1)

        btn_vclink = QPushButton("Vendor Credit Linked Txn")
        btn_vclink.setStyleSheet(self.button_style)
        btn_vclink.clicked.connect(lambda: export_data(fetch_VendorCreditLinkedTxn, "VendorCreditLinkedTxn"))
        grid.addWidget(btn_vclink, 5, 2)

        self.status_label = QLabel("Idle")
        self.status_label.setObjectName("statusLabel")
        grid.addWidget(self.status_label, 6, 0, 1, 3, alignment=Qt.AlignLeft)

        self.tab_export.setLayout(grid)

    # ----- Close handling -----
    def closeEvent(self, event):
        try:
            close_connection()
        finally:
            event.accept()

# ===============================================================
# ---------------------- Globals bound to UI --------------------
# ===============================================================

status_var = _StatusVar()
dsn_var = _DsnVar(DEFAULT_DSN)

# ===============================================================
# ---------------------- Entry Point ----------------------------
# ===============================================================

def main():
    app = QApplication(sys.argv)
    app.setStyleSheet(qdarkstyle.load_stylesheet_pyqt5())

    window = MainWindow()
    window.show()

    # bind globals now that window exists
    status_var.set("Idle")
    dsn_var.set(DEFAULT_DSN)

    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
