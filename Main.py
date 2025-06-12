import pyodbc
import pandas as pd
import os
import win32com.client

# === שלב 1: חיבור למסד הנתונים ויצירת קובץ Excel ===
server = 'sapsrv'
database = 'Civan'
username = 'sa'
password = 'B1Admin'

conn_str = (
    'DRIVER={ODBC Driver 17 for SQL Server};'
    f'SERVER={server};'
    f'DATABASE={database};'
    f'UID={username};'
    f'PWD={password};'
)

# ביצוע השאילתה
conn = pyodbc.connect(conn_str)
query = """
SELECT DocNum, CardName 
FROM OPOR 
WHERE CAST(DocDate AS DATE) = CAST(GETDATE() AS DATE)
"""
df = pd.read_sql(query, conn)
conn.close()

# שמירת התוצאה לקובץ Excel
output_path = os.path.join(os.getcwd(), 'daily_orders.xlsx')
df.to_excel(output_path, index=False, engine='openpyxl')

print(f"הקובץ נשמר בהצלחה בנתיב: {output_path}")

# === שלב 2: שליחת מייל דרך Outlook Desktop ===
outlook = win32com.client.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)

mail.To = "naftali.cohen@civanlasers.com"
mail.Subject = "דו\"ח יומי אוטומטי - הזמנות רכש"
mail.Body = "שלום,\n\nמצורף דו\"ח ההזמנות מהיום.\n\nנשלח אוטומטית מ-Python."

# צרף את הקובץ
mail.Attachments.Add(output_path)

# שלח את המייל
mail.Send()

print("המייל נשלח בהצלחה דרך Outlook Desktop!")
