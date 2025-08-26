import pyodbc

DSN_NAME = "QuickBooks Data"   # apna DSN naam daale jo aapne ODBC me banaya hai

try:
    # connect
    conn = pyodbc.connect(f"DSN={DSN_NAME};", autocommit=True)
    cursor = conn.cursor()

    # ek simple query (Company table Reckon/QuickBooks me hoti hai)
    cursor.execute("SELECT * FROM Company")
    row = cursor.fetchone()

    if row:
        print("✅ Connection Successful!")
        print("Company Info:", row)
    else:
        print("⚠ Connected, but no data returned.")

    conn.close()

except Exception as e:
    print("❌ Connection failed:")
    print(e)