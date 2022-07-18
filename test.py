# # import pandas as pd

# # df = pd.ExcelFile('data.xlsx').parse('Sheet1')
# # a=[]
# # b=[]
# # c=[]
# # d=[]
# # a.append(df['Datum'])
# # b.append(df['Start'])
# # c.append(df['Ende'])
# # d.append(df['Pause'])

# # #erg = [len(a),len(b),len(c),len(d)]

# # #print(erg)

# # print(a)
# #print(b)
# #print(c)
# #print(d)

# # f = open("out.txt", "w")
# # f.write()
# # f.close()


# import openpyxl
# import terminaltables
# table = []
# fname = 'data.xlsx'
# wb = openpyxl.load_workbook(fname)
# sheet = wb.get_sheet_by_name('Sheet1')
# for rowOfCellObjects in sheet['A1':'D274']:
#   for cellObj in rowOfCellObjects:
#     #print(cellObj.coordinate, cellObj.value)
#     table.append([cellObj.coordinate, cellObj.value])

# print(terminaltables.AsciiTable(table,"Stunden").table)