from csv import excel
from excelManager import ExcelManager

excelManager = ExcelManager('./Attendata.xlsx', 'CIE')

print(excelManager.get_student_data(1, 11))
