import xlrd
import xlwt


student_names = []
student_numbers = []

# Opens workbook for reading.
wb = xlrd.open_workbook("excel-sheet.xlsx") 
sheet = wb.sheet_by_index(0) 

# Opens workbook for writing
wb_write = xlwt.Workbook()
ws = wb_write.add_sheet("Admissions List)

# Puts all student names in a list.
for i in range(6, 469, 1):
	student_names.append(sheet.cell_value(i, 1))  

# Puts all student numbers in a list
for k in range(6, 469, 1):
	student_numbers.append(int(sheet.cell_value(k, 2)))
			
# Writes full names and student numbers to second workbook
for j in range(len(student_names)):
	ws.write(j, 1, '{}'.format(student_numbers[j] + ', ' + student_names[j]))
	
ws.save('admissions.xls') # saves file
	
	
