import sys
import pyodbc
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QLabel, QComboBox,
    QPushButton, QDateEdit, QHBoxLayout, QTextEdit, QGroupBox, QFrame
)
from PyQt5.QtCore import QDate, Qt
from PyQt5.QtGui import QFont


class QBOCountsApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Summary Tool")
        self.setGeometry(200, 200, 700, 600)

        main_layout = QVBoxLayout()

        # ---------------- Title ----------------
        title = QLabel("Summary Tool")
        title.setFont(QFont("Arial", 18, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title)

        # ---------------- DSN Section ----------------
        dsn_group = QGroupBox("Database Connection")
        dsn_layout = QVBoxLayout()
        dsn_layout.addWidget(QLabel("Select DSN:"))
        self.dsn_combo = QComboBox()
        self.load_dsns()
        dsn_layout.addWidget(self.dsn_combo)
        dsn_group.setLayout(dsn_layout)
        main_layout.addWidget(dsn_group)

        # ---------------- Date Section ----------------
        date_group = QGroupBox("Date Range")
        date_layout = QHBoxLayout()
        self.from_date = QDateEdit()
        self.from_date.setCalendarPopup(True)
        self.from_date.setDate(QDate.currentDate().addMonths(-1))

        self.to_date = QDateEdit()
        self.to_date.setCalendarPopup(True)
        self.to_date.setDate(QDate.currentDate())

        date_layout.addWidget(QLabel("From:"))
        date_layout.addWidget(self.from_date)
        date_layout.addWidget(QLabel("To:"))
        date_layout.addWidget(self.to_date)
        date_group.setLayout(date_layout)
        main_layout.addWidget(date_group)

        # ---------------- Fetch Button ----------------
        self.fetch_btn = QPushButton("Fetch Counts")
        self.fetch_btn.setStyleSheet("""
            QPushButton {
                background-color: #2E86C1;
                color: white;
                font-size: 14px;
                font-weight: bold;
                padding: 10px 18px;
                border-radius: 8px;
            }
            QPushButton:hover {
                background-color: #1A5276;
            }
        """)
        self.fetch_btn.clicked.connect(self.fetch_counts)
        main_layout.addWidget(self.fetch_btn, alignment=Qt.AlignCenter)

        # ---------------- Separator ----------------
        line = QFrame()
        line.setFrameShape(QFrame.HLine)
        line.setFrameShadow(QFrame.Sunken)
        main_layout.addWidget(line)

        # ---------------- Results Section ----------------
        result_group = QGroupBox("Results")
        result_layout = QVBoxLayout()
        self.result_text = QTextEdit()
        self.result_text.setReadOnly(True)
        self.result_text.setFont(QFont("Courier New", 11))  # Monospace font for alignment
        result_layout.addWidget(self.result_text)
        result_group.setLayout(result_layout)
        main_layout.addWidget(result_group)

        self.setLayout(main_layout)

    def load_dsns(self):
        """Auto fetch system DSNs"""
        dsns = pyodbc.dataSources()
        for dsn in dsns.keys():
            self.dsn_combo.addItem(dsn)

    def fetch_counts(self):
        dsn = self.dsn_combo.currentText()
        from_date = self.from_date.date().toString("yyyy-MM-dd")
        to_date = self.to_date.date().toString("yyyy-MM-dd")

        try:
            conn = pyodbc.connect(f"DSN={dsn}", autocommit=True)
            cursor = conn.cursor()

            # ----------------- Basic Counts -----------------
            cursor.execute("SELECT COUNT(*) FROM Account")
            coa_count = cursor.fetchone()[0]

            cursor.execute("""
                SELECT AccountType, COUNT(*) 
                FROM Account 
                WHERE AccountType IN ('Bank','CreditCard','AccountsPayable','AccountsReceivable')
                GROUP BY AccountType
            """)
            account_type_counts = cursor.fetchall()
            type_counts = {row[0]: row[1] for row in account_type_counts}

            bank_count = type_counts.get("Bank", 0)
            credit_count = type_counts.get("CreditCard", 0)
            ap_count = type_counts.get("AccountsPayable", 0)
            ar_count = type_counts.get("AccountsReceivable", 0)

            cursor.execute("SELECT COUNT(*) FROM Employee")
            emp_count = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM Class")
            class_count = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(*) FROM ItemService")
            item_count = cursor.fetchone()[0]

            # Invoice
            cursor.execute("SELECT COUNT(*) FROM Invoice")
            invoice_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM Invoice
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            invoice_count_date = cursor.fetchone()[0]

            # Bill
            cursor.execute("SELECT COUNT(*) FROM Bill")
            bill_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM Bill
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            bill_count_date = cursor.fetchone()[0]

            # CreditCard
            cursor.execute("SELECT COUNT(*) FROM CreditCardCharge")
            creditCard_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM CreditCardCharge
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            creditCard_count_date = cursor.fetchone()[0]

            # Credit Memo
            cursor.execute("SELECT COUNT(*) FROM CreditMemo")
            CreditMemo_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM CreditMemo
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            CreditMemo_count_date = cursor.fetchone()[0]

            # Sales Receipt
            cursor.execute("SELECT COUNT(*) FROM SalesReceipt")
            SalesReceipt_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM SalesReceipt
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            SalesReceipt_count_date = cursor.fetchone()[0]

            # JournalEntry
            cursor.execute("SELECT COUNT(*) FROM JournalEntry")
            JournalEntry_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM JournalEntry
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            JournalEntry_count_date = cursor.fetchone()[0]

            # VendorCredit
            cursor.execute("SELECT COUNT(*) FROM VendorCredit")
            VendorCredit_count = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*) FROM VendorCredit
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            VendorCredit_count_date = cursor.fetchone()[0]

            # Transaction Total
            cursor.execute("SELECT COUNT(*) FROM Transaction")
            total_line = cursor.fetchone()[0]

            cursor.execute(f"""
                SELECT COUNT(*)
                FROM Transaction
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
            """)
            total_line_with_Date = cursor.fetchone()[0]

            # TxnType counts
            cursor.execute("SELECT TxnType, COUNT(*) FROM Transaction GROUP BY TxnType")
            txn_type_all = cursor.fetchall()

            cursor.execute(f"""
                SELECT TxnType, COUNT(*) 
                FROM Transaction
                WHERE TxnDate >= {{d '{from_date}'}} AND TxnDate <= {{d '{to_date}'}}
                GROUP BY TxnType
            """)
            txn_type_date = cursor.fetchall()

            txn_type_all_dict = {row[0]: row[1] for row in txn_type_all}
            txn_type_date_dict = {row[0]: row[1] for row in txn_type_date}

            # ----------------- Result Text -----------------
            result_text = (
                f"=== First Section ===\n"
                f"COA : {coa_count}\n"
                f"Bank Accounts: {bank_count}\n"
                f"CreditCard Accounts: {credit_count}\n"
                f"AccountsPayable: {ap_count}\n"
                f"AccountsReceivable: {ar_count}\n"
                f"Employee : {emp_count}\n"
                f"Class : {class_count}\n"
                f"Item : {item_count}\n\n"

                f"=== Inception Section ===\n"
                f"Invoice : {invoice_count} (DateRange: {invoice_count_date})\n"
                f"Bill : {bill_count} (DateRange: {bill_count_date})\n"
                f"Credit Card : {creditCard_count} (DateRange: {creditCard_count_date})\n"
                f"Credit Memo : {CreditMemo_count} (DateRange: {CreditMemo_count_date})\n"
                f"Sales Receipt : {SalesReceipt_count} (DateRange: {SalesReceipt_count_date})\n"
                f"Journal Entry : {JournalEntry_count} (DateRange: {JournalEntry_count_date})\n"
                f"Bill Credit : {VendorCredit_count} (DateRange: {VendorCredit_count_date})\n\n"
                f"Total Line : {total_line}\n"
                f"Total Line With Date : {total_line_with_Date}\n\n"
                
                # f"TxnType Counts (All): {txn_type_all_dict}\n\n"
                # f"TxnType Counts (DateRange): {txn_type_date_dict}\n"
                f"*--------------Thank You--------------*"
            )

            self.result_text.setText(result_text)
            conn.close()

        except Exception as e:
            self.result_text.setText(f"Error: {str(e)}")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = QBOCountsApp()
    window.show()
    sys.exit(app.exec_())
