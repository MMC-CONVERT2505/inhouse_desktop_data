import win32com.client
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import threading, pythoncom

# ---------------------- QBXML BUILDERS ---------------------- #
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

def build_invoice_query(version, max_return=20):
    return f"""<?xml version="1.0" encoding="utf-8"?>
<?qbxml version="{version}"?>
<QBXML>
  <QBXMLMsgsRq onError="stopOnError">
    <InvoiceQueryRq requestID="2">
      <MaxReturned>{max_return}</MaxReturned>
      <IncludeLineItems>true</IncludeLineItems>
    </InvoiceQueryRq>
  </QBXMLMsgsRq>
</QBXML>"""

# ---------------------- SDK FUNCTIONS ---------------------- #
def open_connection():
    pythoncom.CoInitialize()  # thread safe init

    # try multiple request processor versions
    rp = None
    for i in range(1, 6):
        try:
            rp = win32com.client.Dispatch(f"QBXMLRP2.RequestProcessor.{i}")
            break
        except:
            continue
    if not rp:
        raise Exception("No compatible RequestProcessor version found.")

    rp.OpenConnection2("", "Reckon Data Fetch Tool", 1)
    ticket = rp.BeginSession("", 2)
    return rp, ticket

def close_connection(rp, ticket):
    rp.EndSession(ticket)
    rp.CloseConnection()
    pythoncom.CoUninitialize()  # thread cleanup

def detect_qbxml_version(rp, ticket):
    try:
        versions = rp.QBXMLVersionsForSession(ticket)
        if isinstance(versions, (list, tuple)):
            version = versions[-1]
        else:
            version = str(versions)
        return version
    except:
        return "6.1"

def fetch_accounts(rp, ticket, version):
    request = build_account_query(version)
    response = rp.ProcessRequest(ticket, request)
    root = ET.fromstring(response)
    return root.findall(".//AccountRet")

def fetch_invoices(rp, ticket, version):
    request = build_invoice_query(version, max_return=100)
    response = rp.ProcessRequest(ticket, request)
    root = ET.fromstring(response)
    return root.findall(".//InvoiceRet")

# ---------------------- EXPORT FUNCTIONS ---------------------- #
def export_to_excel(data, data_type):
    if not data:
        messagebox.showwarning("No Data", f"No {data_type} data to export.")
        return

    try:
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title=f"Save {data_type} Data"
        )
        if not file_path:
            return

        rows = []
        if data_type == "Accounts":
            for acc in data:
                rows.append({
                    "Name": acc.findtext("Name"),
                    "AccountType": acc.findtext("AccountType"),
                    "AccountNumber": acc.findtext("AccountNumber"),
                    "ListID": acc.findtext("ListID"),
                })
        elif data_type == "Invoices":
            for inv in data:
                rows.append({
                    "TxnID": inv.findtext("TxnID"),
                    "RefNumber": inv.findtext("RefNumber"),
                    "Customer": inv.findtext("CustomerRef/FullName"),
                    "TxnDate": inv.findtext("TxnDate"),
                    "BalanceRemaining": inv.findtext("BalanceRemaining"),
                })

        df = pd.DataFrame(rows)
        df.to_excel(file_path, index=False)

        messagebox.showinfo("Export Successful", f"{data_type} exported to:\n{file_path}")

    except Exception as e:
        messagebox.showerror("Export Failed", str(e))

# ---------------------- UI HANDLERS ---------------------- #
last_accounts = []
last_invoices = []

def threaded(fn, log, progress):
    def wrapper():
        progress.start()
        t = threading.Thread(target=lambda: safe_run(fn, log, progress))
        t.start()
    return wrapper

def safe_run(fn, log, progress):
    try:
        fn(log)
    except Exception as e:
        messagebox.showerror("Error", str(e))
        log(f"❌ Error: {e}")
    finally:
        progress.stop()

def handle_fetch_accounts(log):
    global last_accounts
    log("Fetching Chart of Accounts...")

    rp, ticket = open_connection()
    version = detect_qbxml_version(rp, ticket)
    last_accounts = fetch_accounts(rp, ticket, version)
    close_connection(rp, ticket)

    log(f"✅ {len(last_accounts)} accounts fetched!")

def handle_fetch_invoices(log):
    global last_invoices
    log("Fetching Invoices...")

    rp, ticket = open_connection()
    version = detect_qbxml_version(rp, ticket)
    last_invoices = fetch_invoices(rp, ticket, version)
    close_connection(rp, ticket)

    log(f"✅ {len(last_invoices)} invoices fetched!")

def handle_export():
    if last_accounts:
        export_to_excel(last_accounts, "Accounts")
    elif last_invoices:
        export_to_excel(last_invoices, "Invoices")
    else:
        messagebox.showwarning("No Data", "Please fetch COA or Invoices first.")

# ---------------------- MAIN UI ---------------------- #
def main():
    root = tk.Tk()
    root.title("Reckon SDK Data Fetch + Export Tool")
    root.geometry("700x500")
    root.configure(bg="#1e1e1e")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", font=("Segoe UI", 12), padding=10,
                    background="#333", foreground="white", relief="flat")
    style.map("TButton", background=[("active", "#555")])

    lbl = tk.Label(root, text="🧾 Reckon Data Fetch + Export Tool",
                   bg="#1e1e1e", fg="white", font=("Segoe UI", 18, "bold"))
    lbl.pack(pady=10)

    frame = tk.Frame(root, bg="#1e1e1e")
    frame.pack(pady=10)

    progress = ttk.Progressbar(root, mode="indeterminate", length=500)
    progress.pack(pady=5)

    log_box = tk.Text(root, height=12, width=90, bg="#111", fg="white",
                      font=("Consolas", 10), relief="flat")
    log_box.pack(pady=10)

    def log(msg):
        log_box.insert(tk.END, msg + "\n")
        log_box.see(tk.END)

    ttk.Button(frame, text="📂 Fetch COA",
               command=threaded(handle_fetch_accounts, log, progress)).grid(row=0, column=0, padx=15)
    ttk.Button(frame, text="🧾 Fetch Invoices",
               command=threaded(handle_fetch_invoices, log, progress)).grid(row=0, column=1, padx=15)
    ttk.Button(frame, text="📑 Export to Excel",
               command=handle_export).grid(row=0, column=2, padx=15)

    root.mainloop()

if __name__ == "__main__":
    main()
