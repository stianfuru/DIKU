from openpyxl import Workbook 

filename = "testbook.xlsx"

workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "Test123"

sheet["E5"] = "Hallo ja"

sheet["D4"] = "Test43215321"

workbook.save(filename=filename)

#tester pyxl. lager en ny excel fil og skirver inn i sheet