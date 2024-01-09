import sqlite3
import openpyxl

wb = openpyxl.load_workbook(XLFILEPATH)
sheet = wb['sheet1']

conn = sqlite3.connect('data.db')
cursor = conn.cursor()
cursor.execute('''
               CREATE TABLE IF NOT EXISTS student (
                   usn VARCHAR(10),
                   phone_no  INTEGER(10),
                   name TEXT(40)
               ''')


cursor.close()
conn.commit()
conn.close()
