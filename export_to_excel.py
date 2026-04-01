import sqlite3
import pandas as pd

# Connect to your existing database
conn = sqlite3.connect("seoul_office.db")

# List of all your tables
tables = [
    "macro", "existing_supply", "future_supply", "vacancy", 
    "net_absorption", "rent", "capital_value", "cap_rate", "capital_markets"
]

# Create an Excel writer
with pd.ExcelWriter("seoul_office_data.xlsx") as writer:
    for table in tables:
        try:
            # Read each table and save it as a sheet
            df = pd.read_sql(f"SELECT * FROM {table}", conn)
            df.to_excel(writer, sheet_name=table, index=False)
            print(f"✅ Exported {table}")
        except Exception as e:
            print(f"⚠️ Could not export {table}: {e}")

conn.close()
print("🎉 Export complete! You now have 'seoul_office_data.xlsx'")