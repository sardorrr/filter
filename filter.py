import glob
from xlwt import Workbook



wb = Workbook()
sheet1 = wb.add_sheet('Sheet 1')
sheet1.write(0,0,"#")
sheet1.col(0).width = 1500
sheet1.write(0,1,"дата")
sheet1.col(1).width = 1500
sheet1.write(0,2,"время")
sheet1.col(2).width = 1500
sheet1.write(0,3,"Адрес")
sheet1.col(3).width = 7000
sheet1.write(0,4,"оплата")
sheet1.col(4).width = 3000
sheet1.write(0,5,"Итого")
sheet1.col(5).width = 3000

searchfile = open("Cafe.txt", "r")
i = 1;
for line in searchfile:

        #if "#" in line: sheet1.write(i,0, "".join(line.split()[1][1:]))
        if line.find( "Заказ" ) == 0 or line.find( "Заказ" ) == 1:
            sheet1.write(i,0, "".join(line.split()[1][1:]))

        if line.find( "Адрес:" ) == 0:
            sheet1.write(i,3," ".join(line.split()[1:]))
        #if "Метод оплаты:" in line: sheet1.write(i,4," ".join(line.split()[-1]))
        if line.find( "Метод оплаты" ) == 0:
            sheet1.write(i,4, "".join(line.split()[-1]))

        #if "Итого" in line: sheet1.write(i,5," ".join(line.split()[1:-1]))
        if line.find( "Итого" ) == 0:
            sheet1.write(i,5,"".join(line.split()[1:-1]))
            i = i + 1

        #if "Метод оплаты:" in line: i+=1 #incriments i so the next line will be used next

searchfile.close()

wb.save('done.xls')
