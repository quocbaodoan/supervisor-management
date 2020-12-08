import xlrd
import xlwt
import os
import pandas.io.sql as sql
from configparser import ConfigParser
import mysql.connector
import random

list_supervisor = []
list_room = []
list_supervisor_in_room = []
list_supervisor_out_room = []

def readFromExcel(filename):
	excel_data = xlrd.open_workbook(filename)
	sheet_supervisor = excel_data.sheet_by_index(0)
	sheet_room = excel_data.sheet_by_index(1)
	for i in range(1, sheet_supervisor.nrows): 
		pinfo = dict(mgv= int(sheet_supervisor.cell_value(i, 1)), tgv = sheet_supervisor.cell_value(i, 2), dvct = sheet_supervisor.cell_value(i, 4))
		list_supervisor.append(pinfo)

	for i in range(1, sheet_room.nrows): 
		list_room.append(int(sheet_room.cell_value(i, 1)))
	#print("Successfully retrieved all excel data")


def supervisorInRoom():
	for room in list_room:
		pinfo = dict(pt = room)
		gv1 = list_supervisor[random.randint(0, len(list_supervisor) - 1)]
		mgv1 = gv1["mgv"]
		tgv1 = gv1["tgv"]
		dvct1 = gv1["dvct"]
		pinfo.update({"mgv1": mgv1})
		pinfo.update({"tgv1": tgv1})
		pinfo.update({"dvct1": dvct1})
		for supervisor in list_supervisor:
			if (pinfo["mgv1"] == supervisor["mgv"]):
				list_supervisor.remove(supervisor)
				break
				
		gv2 = list_supervisor[random.randint(0, len(list_supervisor) - 1)]
		mgv2 = gv2["mgv"]
		tgv2 = gv2["tgv"]
		dvct2 = gv2["dvct"]
		pinfo.update({"mgv2": mgv2})
		pinfo.update({"tgv2": tgv2})
		pinfo.update({"dvct2": dvct2})
		for supervisor in list_supervisor:
			if (pinfo["mgv2"] == supervisor["mgv"]):
				list_supervisor.remove(supervisor)
				break
		
		list_supervisor_in_room.append(pinfo)

def supervisorOutRoom():
	equal_room_per_outside_supervisor = len(list_room)/len(list_supervisor)
	#print(equal_room_per_outside_supervisor)
	if (equal_room_per_outside_supervisor <= 2):
		for i in range(0, len(list_room), 2):
			gv = list_supervisor[random.randint(0, len(list_supervisor) - 1)]
			pinfo = dict(mgv = gv["mgv"], tgv = gv["tgv"], dvct = gv["dvct"])

			if (i + 1 < len(list_room)):
				pinfo.update(pt = str(list_room[i]) + ", " + str(list_room[i + 1]))
			else:
				pinfo.update(pt = str(list_room[i]))

			for supervisor in list_supervisor:
				if (pinfo["mgv"] == supervisor["mgv"]):
					list_supervisor.remove(supervisor)
					break
			list_supervisor_out_room.append(pinfo)
	else:
		for i in range(0, len(list_supervisor) - 1):
			gv = list_supervisor[random.randint(0, len(list_supervisor) - 1)]
			pinfo = dict(mgv = gv["mgv"], tgv = gv["tgv"], dvct = gv["dvct"])
			count = int(equal_room_per_outside_supervisor)
			total_room = str(list_room[i * count])
			j = 1
			while(count > 1):
				total_room += ", " + str(list_room[i * count + j])
				count -= 1
				j += 1
			pinfo.update(pt = str(total_room))
			for supervisor in list_supervisor:
				if (pinfo["mgv"] == supervisor["mgv"]):
					list_supervisor.remove(supervisor)
					break
			list_supervisor_out_room.append(pinfo)

		gv = list_supervisor[0]
		pinfo = dict(mgv = gv["mgv"], tgv = gv["tgv"], dvct = gv["dvct"])
		count = len(list_room) - len(list_supervisor_out_room) * int(equal_room_per_outside_supervisor)
		total_room = str(list_room[len(list_supervisor_out_room) * int(equal_room_per_outside_supervisor)])
		i = 1
		while(count > 1):
			total_room += ", " + str(list_room[len(list_supervisor_out_room) * int(equal_room_per_outside_supervisor) + i])
			count -= 1
			i += 1
		pinfo.update(pt = str(total_room))
		list_supervisor_out_room.append(pinfo)

	#print(list_supervisor_out_room)


def writeToExcel(filename):
	wb = xlwt.Workbook()

	sheet1 = wb.add_sheet("Giám Thị Phòng Thi")
	sheet1.write(0, 0, "STT")
	sheet1.write(0, 1, "Phòng thi")
	sheet1.write(0, 2, "Mã giám thị 1")
	sheet1.write(0, 3, "Giám thị 1")
	sheet1.write(0, 4, "Đơn vị công tác")
	sheet1.write(0, 5, "Mã giám thị 2")
	sheet1.write(0, 6, "Giám thị 2")
	sheet1.write(0, 7, "Đơn vị công tác")
	for i in range(0, len(list_supervisor_in_room)):
		pinfo = list_supervisor_in_room[i]
		sheet1.write(i + 1, 0, i + 1)
		sheet1.write(i + 1, 1, pinfo["pt"])
		sheet1.write(i + 1, 2, pinfo["mgv1"])
		sheet1.write(i + 1, 3, pinfo["tgv1"])
		sheet1.write(i + 1, 4, pinfo["dvct1"])
		sheet1.write(i + 1, 5, pinfo["mgv2"])
		sheet1.write(i + 1, 6, pinfo["tgv2"])
		sheet1.write(i + 1, 7, pinfo["dvct2"])

	sheet2 = wb.add_sheet("Giám Thị Hành Lang")
	sheet2.write(0, 0, "STT")
	sheet2.write(0, 1, "Mã giám thị")
	sheet2.write(0, 2, "Giám thị")
	sheet2.write(0, 3, "Đơn vị công tác")
	sheet2.write(0, 4, "Phòng giám sát")
	for i in range(0, len(list_supervisor_out_room)):
		pinfo = list_supervisor_out_room[i]
		sheet2.write(i + 1, 0, i + 1)
		sheet2.write(i + 1, 1, pinfo["mgv"])
		sheet2.write(i + 1, 2, pinfo["tgv"])
		sheet2.write(i + 1, 3, pinfo["dvct"])
		sheet2.write(i + 1, 4, pinfo["pt"])

	wb.save(filename)


def connectToDatabase():	
	mydb = mysql.connector.connect(user='root', password='admin',
								host='127.0.0.1',
								database='quanlygiamthi')

	mycursor1= mydb.cursor()
	mycursor1.execute("TRUNCATE TABLE giamthiphongthi")
	mydb.commit()

	mycursor2= mydb.cursor()
	mycursor2.execute("TRUNCATE TABLE giamthihanhlang")
	mydb.commit()

	mycursor3 = mydb.cursor()
	for i in range(0, len(list_supervisor_in_room)):
		pinfo = list_supervisor_in_room[i]
		sql = "INSERT INTO giamthiphongthi (phongthi, magiamthi1, giamthi1, donvicongtac1, magiamthi2, giamthi2, donvicongtac2) VALUES (%s, %s, %s, %s, %s, %s, %s)"
		val = (pinfo["pt"], pinfo["mgv1"], pinfo["tgv1"], pinfo["dvct1"], pinfo["mgv2"], pinfo["tgv2"], pinfo["dvct2"])
		mycursor3.execute(sql, val)
	mydb.commit()

	mycursor4 = mydb.cursor()
	for i in range(0, len(list_supervisor_out_room)):
		pinfo = list_supervisor_out_room[i]
		sql = "INSERT INTO giamthihanhlang (magiamthi, giamthi, donvicongtac, phongthi) VALUES (%s, %s, %s, %s)"
		val = (pinfo["mgv"], pinfo["tgv"], pinfo["dvct"], pinfo["pt"])
		mycursor4.execute(sql, val)
	mydb.commit()