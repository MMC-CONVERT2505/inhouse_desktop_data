import sys
import os
import time
import win32com.client
import xml.etree.ElementTree as ET
import xml.sax.saxutils as saxutils
import pyodbc
import pandas as pd

from PyQt5.QtCore import Qt, QDate
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QTabWidget, QVBoxLayout, QHBoxLayout, QGridLayout, QToolButton,QDateEdit,
    QTextEdit, QPushButton, QProgressBar, QLabel, QComboBox, QFileDialog, QMessageBox, QDialog, QFrame)
import qdarkstyle
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QSizePolicy

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
        log("Connecting with Source...")
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

import pandas as pd

def fetch_data(query, connection, name):
    try:
        return pd.read_sql(query, connection)
    except Exception as e:
        print(f"[ERROR] Failed to fetch {name}: {e}")
        return pd.DataFrame()  # empty DataFrame return karega

def get_date_filter(table_alias=""):
    if not MAIN_WINDOW:
        return ""

    from_date_edit = MAIN_WINDOW.from_date
    to_date_edit = MAIN_WINDOW.to_date

    date_field = f"{table_alias}TxnDate" if table_alias else "TxnDate"

    clauses = []
    if from_date_edit.date() != from_date_edit.minimumDate():
        from_dt = from_date_edit.date().toString("yyyy-MM-dd")
        clauses.append(f"{date_field} >= {{d '{from_dt}'}}")

    if to_date_edit.date() != to_date_edit.minimumDate():
        to_dt = to_date_edit.date().toString("yyyy-MM-dd")
        clauses.append(f"{date_field} <= {{d '{to_dt}'}}")

    if clauses:
        return " WHERE " + " AND ".join(clauses)
    return ""


def fetch_coa(connection): 
    return fetch_data("SELECT * FROM Account", connection, "Chart of Accounts")

def fetch_class(connection): 
    return fetch_data("SELECT * FROM Class", connection, "Class")

def fetch_item(connection): 
    return fetch_data("SELECT * FROM Item", connection, "Item")

def fetch_bill(connection):
    query = "SELECT * FROM Bill"
    query += get_date_filter()
    return fetch_data(query, connection, "Bill")

def fetch_invoice(connection):
    query = "SELECT * FROM Invoice"
    query += get_date_filter()
    return fetch_data(query, connection, "Invoice")

def fetch_receivepayment(connection):
    query = "SELECT * FROM ReceivePaymentLine"
    query += get_date_filter()
    return fetch_data(query, connection, "ReceivePaymentLine")

def fetch_BillPaymentCheckLine(connection):
    query = "SELECT * FROM BillPaymentCheckLine"
    query += get_date_filter()
    return fetch_data(query, connection, "BillPaymentCheckLine")

def fetch_BillPaymentCreditCardLine(connection):
    query = "SELECT * FROM BillPaymentCreditCardLine"
    query += get_date_filter()
    return fetch_data(query, connection, "BillPaymentCreditCardLine")

def fetch_InvoiceLine(connection): 
    query = "SELECT * FROM InvoiceLine"
    query += get_date_filter()
    return fetch_data(query, connection, "InvoiceLine")

def fetch_CreditMemoLinkedTxn(connection):
    query = "SELECT * FROM CreditMemoLinkedTxn"
    query += get_date_filter()
    return fetch_data(query, connection, "CreditMemoLinkedTxn")

def fetch_VendorCreditLinkedTxn(connection):
    query = "SELECT * FROM VendorCreditLinkedTxn"
    query += get_date_filter()
    return fetch_data(query, connection, "VendorCreditLinkedTxn")


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
                "Class": fetch_class(connection),
                "Item": fetch_item(connection),
                "Bill": fetch_bill(connection),
                "Invoice": fetch_invoice(connection),
                "ReceivePaymentLine": fetch_receivepayment(connection),
                "BillPaymentCheckLine": fetch_BillPaymentCheckLine(connection),
                "BillPaymentCreditCardLine": fetch_BillPaymentCreditCardLine(connection),
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
        # self.setWindowTitle("MMC Convert | Reckon , QBD")
        # self.setWindowIcon(QIcon("MMC_Convert.png")) 
        # self.resize(600, 500)
        # Remove default OS title bar
        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowSystemMenuHint)
        
        # --- Custom Title Bar ---
        self.title_bar = QWidget()
        self.title_bar.setStyleSheet("background-color: #222;")
        self.title_layout = QHBoxLayout()
        self.title_layout.setContentsMargins(5, 0, 5, 0)
        self.title_bar.setLayout(self.title_layout)

        # App Icon
        self.icon_label = QLabel()
        self.icon_label.setPixmap(QIcon("MMC_Convert.png").pixmap(24, 24))
        self.title_layout.addWidget(self.icon_label)

        # Window Title
        self.title_label = QLabel("MMC Convert | Reckon , QBD")
        self.title_label.setStyleSheet("color: white; font-weight: bold;")
        self.title_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.title_layout.addWidget(self.title_label)

        # Minimize Button
        btn_min = QPushButton("—")
        btn_min.setFixedSize(24, 24)
        btn_min.setStyleSheet("color: white; background: #444;")
        btn_min.clicked.connect(self.showMinimized)
        self.title_layout.addWidget(btn_min)

        # Maximize/Restore Button
        btn_max = QPushButton("⬜")
        btn_max.setFixedSize(24, 24)
        btn_max.setStyleSheet("color: white; background: #444;")
        btn_max.clicked.connect(self.toggle_max_restore)
        self.title_layout.addWidget(btn_max)

        # Close Button
        btn_close = QPushButton("✕")
        btn_close.setFixedSize(24, 24)
        btn_close.setStyleSheet("color: white; background: #aa2222;")
        btn_close.clicked.connect(self.close)
        self.title_layout.addWidget(btn_close)


        # Tabs
        self.tabs = QTabWidget()

        # Create info button (top-right)
        info_button = QToolButton()
        info_button.setText("ℹ️")
        info_button.setStyleSheet("""
            QToolButton {
                border-radius: 10px;
                background-color: #444;
                color: white;
                font-size: 11px;
                margin: 10px;
                padding: 2px;
                min-width: 20px;
                min-height: 20px;
            }
            QToolButton:hover {
                background-color: #666;
            }
        """)
        info_button.clicked.connect(self.show_info_modal)

        # Layout wrapper for tabs + info button
        self.tabs.setCornerWidget(info_button, Qt.TopRightCorner)
        # self.setCentralWidget(self.tabs)
        self.main_widget = QWidget()
        self.main_layout = QVBoxLayout()
        self.main_layout.setContentsMargins(0,0,0,0)
        self.main_layout.addWidget(self.title_bar)  # custom title bar
        self.main_layout.addWidget(self.tabs)       # your existing tabs
        self.main_widget.setLayout(self.main_layout)
        self.setCentralWidget(self.main_widget)


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
        btn_fetch = QPushButton("📂Connect")
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

        # Styled Label
        label_style = """
            QLabel {
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 12px 20px;
            }
        """

        # Row 0: DSN + Test Button
        lbl_dsn = QLabel("Data Source Name:")
        lbl_dsn.setStyleSheet(label_style)
        grid.addWidget(lbl_dsn, 0, 0)

        self.dsn_combo = QComboBox()
        self.dsn_combo.addItem(DEFAULT_DSN)
        self.dsn_combo.setStyleSheet("font-size: 14px; padding: 8px;")
        grid.addWidget(self.dsn_combo, 0, 1)

        btn_test = QPushButton("Test Connection")
        btn_test.setStyleSheet(self.button_style)
        btn_test.clicked.connect(lambda: connect(self.dsn_combo.currentText()))
        grid.addWidget(btn_test, 0, 2)

        # ---- Separator Line ----
        line = QFrame()
        line.setFrameShape(QFrame.HLine)   # Horizontal line
        line.setFrameShadow(QFrame.Sunken)
        grid.addWidget(line, 1, 0, 1, 3)   # spans 3 columns
        
        lbl_from = QLabel("From Date:")
        lbl_from.setStyleSheet(label_style)
        grid.addWidget(lbl_from, 2, 0)

        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setSpecialValueText("")   # allow blank
        self.from_date.setDateRange(QDate(1900, 1, 1), QDate(2100, 12, 31))
        self.from_date.setDate(self.from_date.minimumDate())  # start as blank
        self.from_date.setStyleSheet("font-size: 14px; padding: 8px;")
        grid.addWidget(self.from_date, 2, 1)

        lbl_to = QLabel("To Date:")
        lbl_to.setStyleSheet(label_style)
        grid.addWidget(lbl_to, 3, 0)

        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setSpecialValueText("")
        self.to_date.setDateRange(QDate(1900, 1, 1), QDate(2100, 12, 31))
        self.to_date.setDate(self.to_date.minimumDate())
        self.to_date.setStyleSheet("font-size: 14px; padding: 8px;")
        grid.addWidget(self.to_date, 3, 1)
        
        # Row 2: Export All Button
        btn_all = QPushButton("Export All Data")
        btn_all.setStyleSheet(self.button_style)
        btn_all.clicked.connect(export_all)
        grid.addWidget(btn_all, 3, 2)

        # Row 3+: Export buttons
        btn_coa = QPushButton("Chart of Accounts")
        btn_coa.setStyleSheet(self.button_style)
        btn_coa.clicked.connect(lambda: export_data(fetch_coa, "ChartOfAccounts"))
        grid.addWidget(btn_coa, 4, 0)

        btn_class = QPushButton("Class")
        btn_class.setStyleSheet(self.button_style)
        btn_class.clicked.connect(lambda: export_data(fetch_class, "Class"))
        grid.addWidget(btn_class, 4, 1)
        
        btn_item = QPushButton("Item")
        btn_item.setStyleSheet(self.button_style)
        btn_item.clicked.connect(lambda: export_data(fetch_item, "Item"))
        grid.addWidget(btn_item, 4, 2)
        
        btn_bill = QPushButton("Bill")
        btn_bill.setStyleSheet(self.button_style)
        btn_bill.clicked.connect(lambda: export_data(fetch_bill, "Bill"))
        grid.addWidget(btn_bill, 5, 0)

        btn_invoice = QPushButton("Invoice")
        btn_invoice.setStyleSheet(self.button_style)
        btn_invoice.clicked.connect(lambda: export_data(fetch_invoice, "Invoice"))
        grid.addWidget(btn_invoice, 5, 1)

        btn_recv = QPushButton("Receive Payment Line")
        btn_recv.setStyleSheet(self.button_style)
        btn_recv.clicked.connect(lambda: export_data(fetch_receivepayment, "ReceivePaymentLine"))
        grid.addWidget(btn_recv, 5, 2)

        btn_bpc = QPushButton("Bill Payment Check Line")
        btn_bpc.setStyleSheet(self.button_style)
        btn_bpc.clicked.connect(lambda: export_data(fetch_BillPaymentCheckLine, "BillPaymentCheckLine"))
        grid.addWidget(btn_bpc, 6, 0)

        btn_bpcc = QPushButton("Bill Payment Credit Card Line")
        btn_bpcc.setStyleSheet(self.button_style)
        btn_bpcc.clicked.connect(lambda: export_data(fetch_BillPaymentCreditCardLine, "BillPaymentCreditCardLine"))
        grid.addWidget(btn_bpcc, 6, 1)
        
        line2 = QFrame()
        line2.setFrameShape(QFrame.HLine)
        line2.setFrameShadow(QFrame.Sunken)
        grid.addWidget(line2, 7, 0, 1, 3)

        btn_invline = QPushButton("Invoice Line")
        btn_invline.setStyleSheet(self.button_style)
        btn_invline.clicked.connect(lambda: export_data(fetch_InvoiceLine, "InvoiceLine"))
        grid.addWidget(btn_invline, 8, 0)

        btn_cmlink = QPushButton("Credit Memo Linked Txn")
        btn_cmlink.setStyleSheet(self.button_style)
        btn_cmlink.clicked.connect(lambda: export_data(fetch_CreditMemoLinkedTxn, "CreditMemoLinkedTxn"))
        grid.addWidget(btn_cmlink, 8, 1)

        btn_vclink = QPushButton("Vendor Credit Linked Txn")
        btn_vclink.setStyleSheet(self.button_style)
        btn_vclink.clicked.connect(lambda: export_data(fetch_VendorCreditLinkedTxn, "VendorCreditLinkedTxn"))
        grid.addWidget(btn_vclink, 8, 2)

        # ---- Separator before status ----
        line3 = QFrame()
        line3.setFrameShape(QFrame.HLine)
        line3.setFrameShadow(QFrame.Sunken)
        grid.addWidget(line3, 9, 0, 1, 3)

        # Status Label (Idle)
        self.status_label = QLabel("Idle...")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setStyleSheet(label_style)
        grid.addWidget(self.status_label, 10, 0, 1, 3, alignment=Qt.AlignLeft)

        self.tab_export.setLayout(grid)

    def show_info_modal(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("About Software")
        layout = QVBoxLayout()

        info_text = """
        <b>Company Name:</b> <u>MMC CONVERT</u><br><br>
        <b>Developed By:</b> <u>SANJAY CHOURASIYA</u><br><br>
        <b>Software:</b> Reckon (2024, 2025), Quickbooks (UK, CA, US)<br><br>
        <b>© 2025 <u>MMC Convert</u></b>. All rights reserved.<br>

        """
        label = QLabel(info_text)
        label.setStyleSheet("color: white; font-size: 13px;")
        layout.addWidget(label)

        btn_ok = QPushButton("Close")
        btn_ok.setStyleSheet("padding: 6px 14px;")
        btn_ok.clicked.connect(dialog.accept)
        layout.addWidget(btn_ok, alignment=Qt.AlignCenter)

        dialog.setLayout(layout)
        dialog.setFixedSize(400, 200)
        dialog.exec_()


    # ----- Close handling -----
    def closeEvent(self, event):
        try:
            close_connection()
        finally:
            event.accept()

    # --- Dragging support ---
        self.old_pos = None

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and event.pos().y() < self.title_bar.height():
            self.old_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.old_pos:
            delta = event.globalPos() - self.old_pos
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.old_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.old_pos = None

    def toggle_max_restore(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()
    
    def toggle_max_restore(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()    

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
