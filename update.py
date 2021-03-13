import gspread
from oauth2client.service_account import ServiceAccountCredentials
from pprint import pprint
import sys
import time

# Auth google sheets
scope = ["https://spreadsheets.google.com/feeds",'https://www.googleapis.com/auth/spreadsheets',"https://www.googleapis.com/auth/drive.file","https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name("credentials.json", scope)
client = gspread.authorize(creds)

# Get spreadsheets
database = client.open("Copy of UCR class difficulty database").sheet1         # The name of the database spreadsheet
responses = client.open("ucr survey (Responses)").sheet1  # The name of the responses spreadsheet

# Use input params as start/end lines
if (len(sys.argv) != 3):
    print("Error - usage is py update.py [startLine] [endLine]")
    print("Example: 'py update.py 4 5' will update based on lines 4 and 5.")
    sys.exit()

startLine = int(sys.argv[1])
endLine = int(sys.argv[2])

print("Starting to process responses spreadsheet on line " + str(startLine) + " and ending on line " + str(endLine) + ".")
total = endLine - startLine + 1
print("Total lines to process is " + str(total))
print("----------------------------------")

# validate input
if (total <= 0):
    print("Error - negative lines to process")
    sys.exit()

for i in range(startLine, endLine + 1):
    # grab line
    line = responses.row_values(i)

    print("processing line " + str(i))
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

    # find location to insert
    cell = database.acell('A1')

    try:
        cell = database.find(respClass, in_column=1)
    except:
        print("Did not find class (" + respClass + ") in Column A. (Does it need to be added?)")
        sys.exit()

    print("Found class response (" + respClass + ") at cell A%s" % (cell.row))

    # get current difficulty
    cellA1Diff = "B" + str(cell.row)
    avgDiff = database.acell(cellA1Diff, value_render_option='FORMULA').value

    print("Updating difficulty")

    # create string for new difficulty
    endParenthesis = avgDiff.rfind(")")
    newAvgDiff = avgDiff[:endParenthesis] + "," + respDiff + avgDiff[endParenthesis:]
    print(str(avgDiff) + " -> " + str(newAvgDiff))

    # update difficulty
    database.update(cellA1Diff, newAvgDiff, value_input_option='USER_ENTERED')
    database.format(cellA1Diff, {
        "numberFormat": {
                "type": "NUMBER",
                "pattern": "##.00"
            }
        })

    # find place to insert review, difficulty, date
    print("Column B: starting at current cell difficulty, finding next difficulty.")

    diffList = database.get(cellA1Diff + ":B" + str(cell.row + 100), value_render_option="FORMULA")  # searches 100 lines below current diff. If more than 100 reviews, this will need to be updated.
    empty_line = "[empty line]"  # update this string into a value that won't be entered as a comment.

    flat_list = []
    for sublist in diffList:
        if len(sublist) == 0:
            flat_list.append(empty_line)
        else:
            for item in sublist:
                flat_list.append(item)

    # remove current difficulty to search for next difficulty
    flat_list.pop(0)

    nextDiff = 0

    for i in range(len(flat_list)):
        if (flat_list[i] != empty_line):
            nextDiff = i
            break
            
    nextDiff += 1
    print("Next difficulty found at cell B" + str(nextDiff + cell.row))

    # add review
    database.insert_row(rowToInsert, index=int(nextDiff + cell.row))
    print("Row insert successful")

    # format newly inserted review
    print("Formatting new row")
    database.format('C' + str(nextDiff + cell.row), {"wrapStrategy":"WRAP"})
    database.format('D' + str(nextDiff + cell.row), {"horizontalAlignment":"CENTER"})
    database.format('E' + str(nextDiff + cell.row), {"horizontalAlignment":"CENTER"})
    database.format('A' + str(nextDiff + cell.row) + ':E' + str(nextDiff + cell.row), {"backgroundColor":{"red":1,"green":1,"blue":1}})

    print("----------------------------------")

print("Success! Added " + str(total) + " reviews.")