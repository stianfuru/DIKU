from openpyxl import Workbook 

filename = "testbook.xlsx"

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Test123"

sheet["F3"] = 'enda en test'

workbook.save(filename=filename)

#tester pyxl. lager en ny excel fil og skirver inn i sheet