import csv, pyodbc


def main():
    # set up some constants
    MDB = 'ddd.mdb'
    DRV = '{Microsoft Access Driver (*.mdb)}'
    PWD = 'pw'

    # connect to db
    con = pyodbc.connect('DRIVER={};DBQ={};PWD={}'.format(DRV,MDB,PWD))
    cur = con.cursor()

    SQL = '''select MSysObjects.name
from MSysObjects
where
   MSysObjects.type In (1,4,6)
   and MSysObjects.name not like '~*'   
   and MSysObjects.name not like 'MSys*'
order by MSysObjects.name'''
    rows = cur.execute(SQL).fetchall()
    print(rows)

    # run a query and get the results
    SQL = 'SELECT * FROM mytable;' # your query goes here
    rows = cur.execute(SQL).fetchall()
    cur.close()
    con.close()
    print(rows)

    # you could change the mode from 'w' to 'a' (append) for any subsequent queries
    with open('mytable.csv', 'w') as fou:
        csv_writer = csv.writer(fou) # default field-delimiter is ","
        csv_writer.writerows(rows)


if __name__ == '__main__':
    main()