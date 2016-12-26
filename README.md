import docxtpl
import xlrd


#Please replace the path with the location of the excel file
path = "/Users/MatthewYoungAir/Desktop/Westeros/westerosmaster.xlsx"
wb = xlrd.open_workbook(path)


#Please replace the template_path with the location of the contactletter template file
template_path = "/Users/MatthewYoungAir/Desktop/Westeros/contactletter.docx"
templateDocument = docxtpl.DocxTemplate(template_path)


#Please replace the path with the location of where you would like the saved letters to go
savehere = "/Users/MatthewYoungAir/Desktop/Westeros/introletter.docx"

savehere = savehere.replace('.docx',"")


sheet = wb.sheet_by_index(0)
column_names = sheet.row_values(0)


for row_index in range(1, sheet.nrows):
    row = sheet.row_values(row_index)
    context = dict(zip(column_names, row))
    templateDocument.render(context)
    templateDocument.save("{}_{}.docx".format(savehere,row_index))
