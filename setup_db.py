import pandas as pd
import sqlite3

# Load Excel sheet
excel_path = 'Copy of NEW  SEMI FINAL ImportData_Transactions -.xlsx'
df = pd.read_excel(excel_path, sheet_name='Assets')

# Connect to database
conn = sqlite3.connect('school_assets.db')
cursor = conn.cursor()

# Drop old tables if re-running
cursor.execute("DROP TABLE IF EXISTS assets")
cursor.execute("DROP TABLE IF EXISTS asset_movements")
cursor.execute("DROP TABLE IF EXISTS disposals")

# Create tables
cursor.execute("""
CREATE TABLE assets (
    asset_code TEXT PRIMARY KEY,
    description TEXT,
    category TEXT,
    asset_group TEXT,
    location TEXT,
    department TEXT,
    asset_user TEXT,
    purchase_date TEXT,
    cost REAL,
    expiry_date TEXT,
    vendor TEXT,
    status TEXT DEFAULT 'Active'
);
""")

cursor.execute("""
CREATE TABLE asset_movements (
    movement_id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_code TEXT,
    from_location TEXT,
    to_location TEXT,
    from_department TEXT,
    to_department TEXT,
    move_date TEXT
);
""")

cursor.execute("""
CREATE TABLE disposals (
    disposal_id INTEGER PRIMARY KEY AUTOINCREMENT,
    asset_code TEXT UNIQUE,
    disposal_date TEXT,
    disposal_cost REAL,
    reason TEXT
);
""")

# Insert records from Excel
for _, row in df.iterrows():
    code = str(row.get("Asset Code", "")).strip()
    if not code or code == "nan":
        continue
    cursor.execute("""
        INSERT OR REPLACE INTO assets 
        (asset_code, description, category, asset_group, location, department, asset_user, purchase_date, cost, expiry_date, vendor, status)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'Active')
    """, (
        code,
        str(row.get("Asset Description", "")).strip(),
        str(row.get("Category", "")).strip(),
        str(row.get("Asset Group", "")).strip(),
        str(row.get("Asset Location", "")).strip(),
        str(row.get("Department", "")).strip(),
        str(row.get("Asset User", "")).strip(),
        str(row.get("Purchase Date(dd/mm/yyyy)", "")).strip(),
        float(row.get("Cost", 0.0) or 0.0),
        str(row.get("Expiry Date(dd/mm/yyyy)", "")).strip(),
        str(row.get("Vendor", "")).strip()
    ))
    import sqlite3

# 1. Update this to the exact path of your SQLite database file
db_path = r"D:\Software\your_database_name.db"  # Change to your actual .db path

conn = sqlite3.connect(db_path)
cursor = conn.cursor()

# Columns required by line 421 in app.py
columns = [
    ("discard_reason", "TEXT"),
    ("discard_date", "TEXT"),
    ("scrap_value", "REAL")
]

for col_name, col_type in columns:
    try:
        cursor.execute(f"ALTER TABLE assets ADD COLUMN {col_name} {col_type};")
        print(f"Added column: {col_name}")
    except sqlite3.OperationalError:
        print(f"Column '{col_name}' already exists.")

conn.commit()
conn.close()

conn.commit()
conn.close()
print("Database school_assets.db created successfully!")