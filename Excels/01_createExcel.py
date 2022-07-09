#pip install xlwt
import xlwt

#create excel
wb = xlwt.Workbook()

#add new sheet
ws = wb.add_sheet('Carb Cycle')

#fill in data ws.write(row,col)
ws.write(0,0,'Intial weight')
ws.write(0,1,'Bone muscle')
ws.write(0,2,'Fat rate')
ws.write(0,3,'Net muscle')
ws.write(0,4,'week')

#save
wb.save('./Data/weight.xls')