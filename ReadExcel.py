from openpyxl import load_workbook

wb = load_workbook(filename = 'xlsx/BMW_Sales_Standards_2016_ME.xlsx', data_only=True)
sheets = wb.sheetnames[11:12] #Current Sheet

# Arrays
numberCategory = []
index_number_categories = []

for sheet in sheets: #In current sheet give me the rows and columns
  ws = wb[sheet] # Pass the info as ws variable
  for row in ws.rows: #Get all rows
    numberCategory.insert(0,row[1].value)  #Get the category number and insert the value of row #1 in excel => B
    number_categories_without_filter = next(i for i in numberCategory if i is not None) #Clean of every None and replacing for the previous valid element

    index_number_categories.insert(0,number_categories_without_filter) # Store data row into a Array to see its index
    print len(index_number_categories)
    # for column in ws.columns:
    #   print column[5].value
  
