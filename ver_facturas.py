import sqlite3
import pandas as pd

conn = sqlite3.connect("facturas.db")
df = pd.read_sql_query("SELECT * FROM facturas ORDER BY fecha_procesamiento DESC", conn)
print(df)
conn.close()