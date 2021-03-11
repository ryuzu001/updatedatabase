import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import sys
import time

# Auth google sheets
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Define global variables
currLine = 0        # current line being processed
totalLines = 0      # total lines to process

# Get spreadsheets
database = client.open("cdbTEST").sheet1
responses = client.open("responsesTEST").sheet1

# Use input params as start/end lines
startLine = int(sys.argv[1])
endLine = int(sys.argv[2])

print("Starting to process responses spreadsheet on line " + str(startLine) + " and ending on line " + str(endLine) + ".")
total = endLine - startLine
print("Total lines to process is " + str(total))
print("----------------------------------")

# validate input
if (total <= 0):
    print("Error - negative lines to process")
    sys.exit()

# grab line
line = responses.row_values(startLine)

# separate out different answers
respDate = line[0]
respClass = line[1]
respDiff = line[2]
respComment = line[3]

# shave timestamp off date
respDateSpace = respDate.find(' ')
respDate = respDate[:respDateSpace]

# build row to insert in database
rowToInsert = ["", "", respComment, respDiff, respDate]

print("Inserting the following row into the database: ")
print(rowToInsert)

# find location to insert
cell = database.find(respClass)
if (cell.col != 1):
    print("Error - " + cell.col + " not found in column 1.")
    sys.exit()

print("Found class response (" + respClass + ") at Row %s Column %s" % (cell.row, cell.col))

cellA1Diff = "B" + str(cell.row)

print(cellA1Diff)

print(database.acell(cellA1Diff, value_render_option='FORMULA').value)



# data = sheet.get_all_records()  # Get a list of all records

# row = sheet.row_values(3)  # Get a specific row
# col = sheet.col_values(3)  # Get a specific column
# cell = sheet.cell(1,3).value  # Get the value of a specific cell

# pprint(data)

# print(col)

# insertRow = ["hello", 5, "red", "blue"]
# sheet.insert_row(insertRow, 4)  # Insert the list as a row at index 4
# sheet.delete_rows(4)

# sheet.update_cell(8,8, "haha")  # Update one cell

# numRows = sheet.row_count  # Get the number of rows in the sheet