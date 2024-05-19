# updatedatabase
Update the UCR Class Database from provided responses automatically

____________

This script was created to help maintain the [UCR Class Difficulty Database](https://docs.google.com/spreadsheets/d/1qiy_Oi8aFiPmL4QSTR3zHe74kmvc6e_159L1mAUUlU0/). 
Students provide responses in the form of a Google Sheet, and until now I have been manually updating the database.
This has proven inefficient as the number of responses gradually rose through the years, and the time has come to automate the process as much as possible.

------------------

This is what the database looks like:

![Database Image](https://github.com/ryuzu001/updatedatabase/blob/master/images/database.PNG)

And here is what the response sheet looks like:

![Response Sheet](https://github.com/ryuzu001/updatedatabase/blob/master/images/responses.PNG)

As you can see, manually updating each response can be time consuming, especially at the end of the quarter with possibly 100+ responses.
So as a result, I have created a script to automatically update the database from the sheet responses.

--------------------------

# Usage

First, install [python](https://www.python.org/) on your computer

Second, install [gspread](https://gspread.readthedocs.io/en/latest/)

Next, create a new project in Google Cloud Platform, and enable the Google Drive and Google Sheets API. Download the `.json` file.
[Here (first 3 minutes)](https://www.youtube.com/watch?v=cnPlKLEGR7E) is a nice video tutorial.

Once your project is set up, create a Google Cloud Service Account. This account is the account that the script will use to update the sheet itself. Download the Service Account's private key and name it `credentials.json`. Put this file in the same level as the `update.py` script.

Finally, navigate to the location of `update.py` and run it. The first parameter is the line in the responses spreadsheet to start at, and the second is the line to end at (inclusive)

`py update.py 400 500` will look at the responses spreadsheet lines 400-500 inclusive and update the database with those reviews.

![Example Image](https://github.com/ryuzu001/updatedatabase/blob/master/images/running.PNG)

------------

# Additional Info

Some of the values in the script are hard coded, and would need an update if you wanted to move things around (Ie, put date on the left-most column)
The script also checks for an empy cell, and assigns it a value that it knows a user would not enter. If they enter this value the script will mess up.
New classes must be added manually. `update.py` will not add new classes to the spreadsheet.
Finally, the script will do it's best to match user-input classes to the values in the spreadsheet, (Ie, the script recognizes 'cs 100' as 'CS100'),
but it can fail, and it is on the updater to go into the responses and update the name. It is also on the user to remove any obvious duplicate answers from the responses as well.


