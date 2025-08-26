# Special Char COA Change

import win32com.client
import xml.etree.ElementTree as ET
import tkinter as tk
from tkinter import ttk, messagebox
import time
import xml.sax.saxutils as saxutils

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

def build_account_mod(version, listid, edit_seq, name, acct_type, new_number):
    safe_name = saxutils.escape(name)  # Escape XML special characters
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

# ---------------------- SDK FUNCTIONS ---------------------- #
def open_connection():
    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    rp.OpenConnection2("", "Reckon COA Tool", 1)
    ticket = rp.BeginSession("", 2)
    return rp, ticket

def close_connection(rp, ticket):
    rp.EndSession(ticket)
    rp.CloseConnection()

def detect_qbxml_version(rp, ticket):
    try:
        versions = rp.QBXMLVersionsForSession(ticket)
        if isinstance(versions, (list, tuple)):
            version = versions[-1]
        else:
            version = str(versions)
        print(f"✅ Detected QBXML Version: {version}")
        return version
    except Exception as e:
        print(f"❌ Version detection failed: {e}")
        return "6.1"

def fetch_accounts(rp, ticket, version):
    request = build_account_query(version)
    response = rp.ProcessRequest(ticket, request)
    root = ET.fromstring(response)
    accounts = root.findall(".//AccountRet")
    return accounts

def auto_numbering(rp, ticket, version, log_callback):
    accounts = fetch_accounts(rp, ticket, version)

    # Collect existing numbers
    used_numbers = []
    name_set = set()  # To handle duplicate names
    for acc in accounts:
        num = acc.findtext("AccountNumber")
        name = acc.findtext("Name")
        if num and num.isdigit():
            used_numbers.append(int(num))
        if name:
            name_set.add(name)

    next_number = max(used_numbers)+1 if used_numbers else 1000
    updated_count = 0

    for acc in accounts:
        listid = acc.findtext("ListID")
        edit_seq = acc.findtext("EditSequence")
        name = acc.findtext("Name")
        acct_type = acc.findtext("AccountType")
        num = acc.findtext("AccountNumber")

        if not num:
            try:
                if not edit_seq:
                    msg = f"⚠ Skipped {name}: EditSequence missing"
                    log_callback(msg)
                    print(msg)
                    continue

                # Handle duplicate names by appending a suffix
                safe_name = name
                suffix = 1
                while safe_name in name_set:
                    suffix += 1
                    safe_name = f"{name} ({suffix})"
                name_set.add(safe_name)

                rp.ProcessRequest(ticket, build_account_mod(version, listid, edit_seq, safe_name, acct_type, str(next_number)))
                msg = f"✔ {name} → Assigned #{next_number}"
                log_callback(msg)
                print(msg)
                next_number += 1
                updated_count += 1

                time.sleep(0.05)  # Small delay to avoid SDK overload

            except Exception as e:
                msg = f"❌ Error updating {name}: {e}"
                log_callback(msg)
                print(msg)

    return updated_count

# ---------------------- UI FUNCTIONS ---------------------- #
def handle_fetch(log, progress):
    try:
        progress.start()
        log("Fetching Chart of Accounts...")
        print("Fetching Chart of Accounts...")

        rp, ticket = open_connection()
        version = detect_qbxml_version(rp, ticket)

        accounts = fetch_accounts(rp, ticket, version)
        close_connection(rp, ticket)

        log(f"✅ {len(accounts)} accounts fetched!")
        print(f"✅ {len(accounts)} accounts fetched!")
        for acc in accounts[:20]:
            log(f" - {acc.findtext('Name')}")
        if len(accounts) > 20:
            log("...")

    except Exception as e:
        log(f"❌ Error: {e}")
        print(f"❌ Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        progress.stop()

def handle_auto_number(log, progress):
    try:
        progress.start()
        log("Running Auto Numbering...")
        print("Running Auto Numbering...")

        rp, ticket = open_connection()
        version = detect_qbxml_version(rp, ticket)

        updated = auto_numbering(rp, ticket, version, log)
        close_connection(rp, ticket)

        log(f"✅ {updated} accounts updated with missing numbers!")
        print(f"✅ {updated} accounts updated with missing numbers!")
        messagebox.showinfo("Auto Numbering", f"{updated} accounts updated successfully!")

    except Exception as e:
        log(f"❌ Error: {e}")
        print(f"❌ Error: {e}")
        messagebox.showerror("Error", str(e))
    finally:
        progress.stop()

# ---------------------- MAIN UI ---------------------- #
def main():
    root = tk.Tk()
    root.title("Reckon COA Numbering Tool")
    root.geometry("600x400")
    root.configure(bg="#1e1e1e")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("TButton", font=("Segoe UI", 12), padding=10,
                    background="#333", foreground="white", relief="flat")
    style.map("TButton", background=[("active", "#555")])

    lbl = tk.Label(root, text="🧾 Reckon COA Manager", bg="#1e1e1e",
                   fg="white", font=("Segoe UI", 18, "bold"))
    lbl.pack(pady=10)

    frame = tk.Frame(root, bg="#1e1e1e")
    frame.pack(pady=10)

    progress = ttk.Progressbar(root, mode="indeterminate", length=400)
    progress.pack(pady=5)

    log_box = tk.Text(root, height=10, width=70, bg="#111", fg="white",
                      font=("Consolas", 10), relief="flat")
    log_box.pack(pady=10)

    def log(msg):
        log_box.insert(tk.END, msg + "\n")
        log_box.see(tk.END)

    ttk.Button(frame, text="📂 Fetch COA", command=lambda: handle_fetch(log, progress)).grid(row=0, column=0, padx=15)
    ttk.Button(frame, text="🔢 Auto Number Missing", command=lambda: handle_auto_number(log, progress)).grid(row=0, column=1, padx=15)

    root.mainloop()

if __name__ == "__main__":
    main()
