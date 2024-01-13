from csv import excel
from excelManager import ExcelManager

excelManager = ExcelManager('./Attendata.xlsx', 'CIE')


for i in range(0, 13):
    print(excelManager.get_student_data(1, i))
