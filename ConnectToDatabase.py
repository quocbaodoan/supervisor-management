import mysql.connector

mydb = mysql.connector.connect(user='root', password='admin',
                            host='127.0.0.1',
                            database='quanlygiamthi')

print(mydb)

mycursor = mydb.cursor()

sql = "INSERT INTO supervisorinroom (room, idsupervisor1, supervisor1, supervisorworkunit1, idsupervisor2, supervisor2, supervisorworkunit2) VALUES (%s, %s, %s, %s, %s, %s, %s)"
val = ("301", "Highway 21", "Highway 21", "Highway 21", "Highway 21", "Highway 21", "Highway 21")
mycursor.execute(sql, val)

mydb.commit()
