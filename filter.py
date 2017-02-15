import glob
from datetime import datetime
from xlwt import Workbook

wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0,0,"#")
sheet1.col(0).width = 1500
sheet1.write(0,1,"дата")
sheet1.col(1).width = 2500
sheet1.write(0,2,"время")
sheet1.col(2).width = 1500
sheet1.write(0,3,"Имя")
sheet1.col(3).width = 5500
sheet1.write(0,4,"телефон")
sheet1.col(4).width = 3500
sheet1.write(0,6,"Филиал")
sheet1.col(6).width = 4000
sheet1.write(0,5,"Адрес")
sheet1.col(5).width = 7000
sheet1.write(0,7,"оплата")
sheet1.col(7).width = 3000
sheet1.write(0,8,"Итого")
sheet1.col(8).width = 2500


with open("Cafe.txt", 'r') as f:
    orderList = [line.split('\n') for line in f.readlines()]
    i = 1
    for m in range(len(orderList)):
        if "Заказ" in str(orderList[m]):
            date = str(orderList[m-1][-2])
            sheet1.write(i,1, date.split()[-2][1:])
            sheet1.write(i,2, date.split()[-1][:-1])
            i +=1



searchfile = open("Cafe.txt", "r")
i = 1;
for line in searchfile:

        #if "#" in line: sheet1.write(i,0, "".join(line.split()[1][1:]))
        if line.find( "Заказ" ) == 0 or line.find( "Заказ" ) == 1:
            sheet1.write(i,0, float("".join(line.split()[1][1:])))

        if line.find( "Имя:" ) == 0:
            sheet1.write(i,3, "".join(line.split()[1:]))

        if line.find( "Телефон:" ) == 0:
            if "+" in line:
                sheet1.write(i,4,("".join(line.split()[1][1:])))
            else:
                sheet1.write(i,4,("".join(line.split()[1:])))

        if line.find( "Филиал: " ) == 0:
            sheet1.write(i,6, "".join(line.split()[1:]))

        if line.find( "Адрес:" ) == 0:
            if " ?? " in line:
                sheet1.write(i,5," ".join(line.split()[2:]))
            else:
                sheet1.write(i,5," ".join(line.split()[1:]))

        #if "Метод оплаты:" in line: sheet1.write(i,4," ".join(line.split()[-1]))
        if line.find( "Метод оплаты" ) == 0:
            sheet1.write(i,7, "".join(line.split()[-1]))

        #if "Итого" in line: sheet1.write(i,5," ".join(line.split()[1:-1]))
        if line.find( "Итого" ) == 0:
            sheet1.write(i,8,float("".join(line.split()[1:-1])))
            i = i + 1

        #if "Метод оплаты:" in line: i+=1 #incriments i so the next line will be used next

searchfile.close()

wb.save('%s.xls'%datetime.now().strftime("%m-%d-%Y___%H-%M-%S"))
