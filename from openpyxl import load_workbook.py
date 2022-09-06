from openpyxl import load_workbook

#load excel file
workbook = load_workbook(filename="C:\Users\AS\OneDrive\Documents\Excel Assingment\Book1.xlsx")

#open workbook
sheet = workbook.active

#modify the desired cell
sheet["B2"] = "Name"

#save the file
workbook.save(filename="C:\Users\AS\OneDrive\Documents\Excel Assingment\output.xlsx")


