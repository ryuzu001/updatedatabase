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
# database = client.open("UCR class difficulty database").sheet1         # The name of the database spreadsheet
# responses = client.open("ucr survey (Responses)").sheet1               # The name of the responses spreadsheet
database = client.open("cdbTEST").sheet1         # The name of the database spreadsheet
responses = client.open("responsesTEST").sheet1               # The name of the responses spreadsheet

# Use input params as start/end lines
if (len(sys.argv) != 3):
    print("Error - usage is py update.py [startLine] [endLine]")
    print("Example: 'py update.py 4 5' will update based on lines 4 and 5.")
    sys.exit()

startLine = 0
endLine = 0
try:
    startLine = int(sys.argv[1])
    endLine = int(sys.argv[2])
except:
    print("Please enter valid parameters.")
    print("Example: 'py update.py 4 5' will update based on lines 4 and 5.")
    sys.exit()

print("Starting to process responses spreadsheet on line " + str(startLine) + " and ending on line " + str(endLine) + ".")
total = endLine - startLine + 1
print("Total lines to process is " + str(total))
print("----------------------------------")

# validate input
if (total <= 0):
    print("Error - negative lines to process")
    sys.exit()

# all input validated, loop through lines and update
for i in range(startLine, endLine + 1):

    existingClass = 1

    # grab line
    line = []
    try:
        line = responses.row_values(i)
    except:
        print("Empty line found at " + str(i) + ". Please try again with valid line numbers.")
        sys.exit()

    # separate out different answers
    respDate = line[0]
    respClass = line[1]
    respDiff = line[2]

    # format class name to upper, remove spaces
    respClass = respClass.upper()
    respClass = respClass.replace(" ", "")

    # shave timestamp off date
    respDateSpace = respDate.find(' ')
    respDate = respDate[:respDateSpace]

    print("processing line " + str(i))
    pprint(line)
    print()

    # allow time for user to review line, and help prevent APIError 429 RESOURCE_EXHAUSTED.
    time.sleep(2)

    # find location to class
    cell = database.acell('A1')

    # if class exists, use existing review, else ask for new review row
    try:
        cell = database.find(respClass, in_column=1)
    except:
        existingClass = 0
        print("Did not find class (" + respClass + ") in Column A.")
        newRow = input("To add new class, please enter row number to use, else press enter: ")
        if (not newRow.isnumeric()):
            print("sorry not a number bro")
            sys.exit()

    if (existingClass):
        print("Found class response (" + respClass + ") at cell A%s" % (cell.row))

        # get current difficulty
        cellA1Diff = "B" + str(cell.row)
        avgDiff = database.acell(cellA1Diff, value_render_option='FORMULA').value

        # class is listed under 2 names (ie, CS111 and MATH111), and the difficulty reads 'See [class]'
        if (avgDiff[:3] == "See"):
            # Find the matching class
            cell = database.find(avgDiff[4:], in_column=1)
            print("Updating class response to " + cell.value + ". Cell updated to A%s" % (cell.row))
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

        # handle reviews with no comment
        if (len(line) != 4):
            print("No review found, skipping add review")
        else:
            # Review has a comment, get comment        
            respComment = line[3]

            # build row to insert in database
            rowToInsert = ["", "", respComment, respDiff, respDate]

            # find place to insert review, difficulty, date
            print("Finding next difficulty...")

            # searches 100 lines below current diff. If more than 100 reviews, this will need to be updated.
            diffList = database.get(cellA1Diff + ":B" + str(cell.row + 100), value_render_option="FORMULA")  
            empty_line = "[empty string]"  # update this string into a value that won't be entered as a comment.

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

            if (nextDiff == 1 and str(database.acell("C" + str(int(cell.row))).value) == "None"):
                # currently no comments, but has a difficulty

                cell_list = database.range('C' + str(cell.row) + ':E' + str(cell.row))
                cell_list[0].value = respComment
                cell_list[1].value = respDiff
                cell_list[2].value = respDate

                database.update_cells(cell_list)
                print("Added first review for existing class.")
            else:
                # check for 'Note: Formerly [class]' 
                formerCheck = str(database.acell('C' + str(int(nextDiff + cell.row - 1))).value)
                if (formerCheck[:14] == "Note: Formerly"):
                    nextDiff -= 1

                # add row
                database.insert_row(rowToInsert, index=int(nextDiff + cell.row))
                print("Row insert successful")

            database.format('A' + str(nextDiff + cell.row) + ':E' + str(nextDiff + cell.row), {"backgroundColor":{"red":1,"green":1,"blue":1}})
        
    else:
        print("Inserting NEW review for class " + respClass + " on row " + newRow)

        # no review
        if (len(line) != 4):
            rowToInsert = [respClass, '=average(' + respDiff + ')', "", "", ""]
        else:
            respComment = line[3]
            rowToInsert = [respClass, '=average(' + respDiff + ')', respComment, respDiff, respDate]

        # insert new row
        database.insert_row(rowToInsert, index=int(newRow), value_input_option='USER_ENTERED')

        # insert_row value_input_option does not seem to work... workaround
        database.update('B' + newRow, '=average(' + respDiff + ')', value_input_option='USER_ENTERED')

    print("----------------------------------")

    # allow time for user to review
    time.sleep(2)

# format all rows
print("Formatting all rows")

# Wrap column C
database.format('C', {"wrapStrategy":"WRAP"})

# Center align columns D and E
database.format('D:E', {"horizontalAlignment":"CENTER"})

print("Success! Added " + str(total) + " review(s).")
