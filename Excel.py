from openpyxl import Workbook 

filename = "testbook.xlsx"

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Test123"

sheet["D4"] = "Test43215321"

sheet["E5"] = 'yesyes'

workbook.save(filename=filename)

#tester pyxl. lager en ny excel fil og skirver inn i sheet