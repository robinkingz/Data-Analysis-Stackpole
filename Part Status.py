import xlrd
location=('C:/Users/robin/Desktop/Learning/Python/S1141.xlsx')
wb = xlrd.open_workbook(location)
sheet = wb.sheet_by_index(0)
Total_reject_count=0
Total_run=0
#print(Total_reject_count)
#print(sheet.cell_value(3,4))
class Part:
    Type_Name=''
    Good=0
    Reject=0
    Ran=False

Part1=Part()
Part2=Part()
Part3=Part()
Part4=Part()
Part5=Part()
for x in range(sheet.nrows):
    if not(isinstance(sheet.cell_value(x,3), str)):
       Total_run=Total_run+1 
       if sheet.cell_value(x,3)!=11410:
          Total_reject_count = Total_reject_count + 1
        #print(sheet.cell_value(x,3))
    if isinstance(sheet.cell_value(x,1),str) and (sheet.cell_value(x,1)!='') and (sheet.cell_value(x,1)[0] == '4'):
       Part1.Ran=True
       Part1.Type_Name='CSS'
       if sheet.cell_value(x,3)!=11410:
          Part1.Reject=Part1.Reject+1
    if isinstance(sheet.cell_value(x,1),str) and (sheet.cell_value(x,1)!='') and (sheet.cell_value(x, 1)[0] == '1') and (sheet.cell_value(x, 1)[15] == 'B'):
       Part2.Ran = True
       Part2.Type_Name = '10R60'
       if sheet.cell_value(x, 3) != 11410:
          Part2.Reject = Part2.Reject + 1
    if isinstance(sheet.cell_value(x,1),str) and (sheet.cell_value(x,1)!='') and (sheet.cell_value(x, 1)[0] == '1') and (sheet.cell_value(x, 1)[15] == 'C'):
       Part2.Ran = True
       Part2.Type_Name = 'AB1V'
       if sheet.cell_value(x, 3) != 11410:
          Part2.Reject = Part2.Reject + 1
print('Total run is',Total_run)
print('Total reject is',Total_reject_count)
print(Part1.Type_Name,Part1.Reject,Part2.Type_Name,Part2.Reject,Part3.Type_Name,Part3.Reject,Part4.Type_Name,Part4.Reject)



#print (Total_reject_count)
#t=type (sheet.cell_value(1,2))
#print (t)
#print(sheet.cell_value(4,1)[0])
#t=type (sheet.cell_value(4,1))
#print (t)
# For row 0 and column 0
#print(part_num)'''