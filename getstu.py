import sqlite3
import csv

con = sqlite3.connect('work17.db')
cur = con.cursor()

idlst = []
with open('idcode.csv') as csvf:
    mycsvreader = csv.reader(csvf)
    for row in mycsvreader:
        idlst.append(row)
print(idlst)

for idcode in idlst:

    print(idcode)
    cur.execute("select count(*) from stus where name=?",idcode)
    res = cur.fetchone()[0]
    if res <= 0:
        print("Do not find!")
    elif res > 1:
        print("Find more stus!")
    else:
        cur.execute("insert into sign (name,idcode,tel,addr,school) select name,idcode,tel,addr,srcs from stus where name=?",idcode)
        con.commit()
cur.close()
con.close()
