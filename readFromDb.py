import csv, pyodbc

print([x for x in pyodbc.drivers()])


# set up some constants
MDB = 'G:\Mijn Drive\ExcelGenerator\20240203-KAZSC.mdb'
DRV = '{Microsoft Access Driver (*.mdb)}'

# connect to db
con = pyodbc.connect('DRIVER={};DBQ={}'.format(DRV,MDB))
cur = con.cursor()

# run a query and get the results 
SQL = 'SELECT * FROM members;' # your query goes here
rows = cur.execute(SQL).fetchall()
cur.close()
con.close()

# you could change the mode from 'w' to 'a' (append) for any subsequent queries
with open('mytable.csv', 'w') as fou:
    csv_writer = csv.writer(fou) # default field-delimiter is ","
    csv_writer.writerows(rows)