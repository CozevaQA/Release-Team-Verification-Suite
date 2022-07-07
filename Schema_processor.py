import openpyxl
from openpyxl import Workbook

file_location = "assets/SchemaForAllWorksheets.xlsx"

Excel_file = openpyxl.load_workbook(file_location)

# Get workbook active sheet object
# from the active attribute
Excel_sheet = Excel_file.active

def getCurrentSchema():
    currentSchema = []
    row_counter = 2
    currentCustID = str(Excel_sheet['A1'].value)
    while currentCustID != 'None':
        current_row = Excel_sheet[row_counter]
        temp = []
        temp.append(str(current_row[0].value))
        temp.append(str(current_row[1].value))
        temp.append(str(current_row[2].value))
        temp.append(str(current_row[3].value))
        temp.append(str(current_row[4].value))
        temp.append(str(current_row[5].value))
        temp.append(str(current_row[6].value))
        currentSchema.append(temp)
        temp = []

        #currentSchema.append(Excel_sheet[row_counter])
        currentCustID = str(Excel_sheet['A'+str(row_counter+1)].value)
        row_counter+=1


    print(currentSchema)
    return currentSchema

def loadSchema(dda):
    wb = Workbook()
    temp_list = []
    wb.save("assets"+ "\\SchemaForAllWorksheets.xlsx")
    ws = wb.active
    ws.append(["CustID","NAME","YEAR","	MEDICARE","COMMERCIAL","UTILIZATION USAGE" ])
    for x in dda:
        for i in range(len(x)):
            temp_list.append(x[i])
        ws.append(temp_list)
        temp_list.clear()

    wb.save("assets" + "\\SchemaForAllWorksheets.xlsx")


#getCurrentSchema()



