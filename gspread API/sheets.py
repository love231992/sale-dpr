import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint

scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]

creds = ServiceAccountCredentials.from_json_keyfile_name("creds.json", scope)

client = gspread.authorize(creds)

sheet = client.open("Tutorial").sheet1                       # Open the spreadhseet 
data = sheet.get_all_records()                               # Get a list of all records
row = sheet.row_values(1)                                    # Get a specific row
col = sheet.col_values(2)                                    # Get a specific column
cell = sheet.cell(3,3).value                                 # Get the value of a specific cell

insertRow = [6,"Ambuja Nalagarh",14]
# sheet.insert_row(insertRow, 8)                             # Insert a row              
# sheet.update_cell(2,2, "CHANGED")                          # Change the value of particular cell
# sheet.update_title("Sheet1")                                 # Update the sheet name
# sheet.delete_row(4)                                          # delete the specific row
numRows = sheet.row_count                                    # count the num of rows in the sheet
len(data)                                                    # count the num of row having data
# pprint(len(data))

rows = [['Asian Fine', '499.68', 7], ['UTCL Bagheri', '488.14', 7], ['ACC LTD', '337.25', 10], ['Fateh', '177.38', 5], ['Everest', '64.17', 1], ['Hemkund Sahib', '191.96', 5], ['Rakesh kumar', '87.49', 3], ['Amritsaria', '97.78', 3], ['Jai Shiv shankar', '85.96', 3], ['Paras', '34.13', 1], ['BTS', '30.17', 1], ['S.A.Bricks', '30.53', 1], ['M.M. Concrete', '28.76', 1], ['Aniket', '32.92', 1], ['Total', '2215.2', 50]]
print(type(rows[0][2]))


# for pos,i in enumerate(rows):
#     sheet.insert_row(rows[pos], pos+1)

