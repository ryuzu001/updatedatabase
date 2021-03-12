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
if (len(sys.argv) != 3):
    print("Error - usage is py update.py [startLine] [endLine]")
    sys.exit()

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

print("processing line " + str(startLine))
pprint(line)
print()

# separate out different answers
respDate = line[0]
respClass = line[1]
respDiff = line[2]
respComment = line[3]

# format class name to upper, remove spaces
respClass = respClass.upper()
respClass = respClass.replace(" ", "")

# shave timestamp off date
respDateSpace = respDate.find(' ')
respDate = respDate[:respDateSpace]

# build row to insert in database
rowToInsert = ["", "", respComment, respDiff, respDate]

print("Inserting the following row into the database: ")
print(rowToInsert)

# find location to insert
cell = database.acell('A1')

try:
    cell = database.find(respClass, in_column=1)
except:
    print("Did not find class (" + respClass + ") in Column A. (Does it need to be added?)")
    sys.exit()

print("Found class response (" + respClass + ") at Row %s Column %s" % (cell.row, cell.col))

# get current difficulty
cellA1Diff = "B" + str(cell.row)
avgDiff = database.acell(cellA1Diff, value_render_option='FORMULA').value

print(avgDiff)
print("Updating difficulty")

# create string for new difficulty
endParenthesis = avgDiff.rfind(")")
newAvgDiff = avgDiff[:endParenthesis] + "," + respDiff + avgDiff[endParenthesis:]
print(newAvgDiff)

# update difficulty
database.update(cellA1Diff, newAvgDiff, value_input_option='USER_ENTERED')
database.format(cellA1Diff, {
    "numberFormat": {
            "type": "NUMBER",
            "pattern": "##.00"
        }
    })

# find place to insert review, difficulty, date

print()
print("Column B, starting at current cell difficulty, finding next difficulty.")
print()

testVar = database.get(cellA1Diff + ":B" + str(cell.row + 90), value_render_option="FORMULA")

flat_list = []
for sublist in testVar:
    if len(sublist) == 0:
        flat_list.append("[empty line]")
    else:
        for item in sublist:
            flat_list.append(item)

for i in range(len(flat_list)):
    print(flat_list[i])



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