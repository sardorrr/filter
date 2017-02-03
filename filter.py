import glob
from xlwt import Workbook


#creates a excel file
wb = Workbook()
#creates new sheet named 'Sheet 1'
sheet1 = wb.add_sheet('Sheet 1')
#Gives column names "#", "дата", "время"....
sheet1.write(0,0,"#")
#Sets column width to 1500
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

#opens Cafe.txt in read only mode
searchfile = open("Cafe.txt", "r")
i = 1;
for line in searchfile:

        #if it finds string "Заказ" in line at position [0] or [1] then it takes the numbers after string "Заказ" and copies it to exel
        if line.find( "Заказ" ) == 0 or line.find( "Заказ" ) == 1:
            sheet1.write(i,0, "".join(line.split()[1][1:]))

        if line.find( "Адрес:" ) == 0:
            sheet1.write(i,3," ".join(line.split()[1:]))

        if line.find( "Метод оплаты" ) == 0:
            sheet1.write(i,4, "".join(line.split()[-1]))

        if line.find( "Итого" ) == 0:
            sheet1.write(i,5,"".join(line.split()[1:-1]))
            i = i + 1 #incriments i so the next row will be used to write data in exel

        

searchfile.close()

wb.save('done.xls')
