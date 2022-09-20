## Koffie Insurance System/Reporting Engineer - Coding Challenge

### Installing dependencies 
To install the Python application dependencies run the following command.
- pip3 install -r requirements.txt

### How to run Python application 
Run the following command
- python3 main.py

### Test cases
I created test cases using the unit test framework.
To run the test case run the following command while being outside the tests folder and in the same directory as main.py
- python3 tests/test_.py

### Notes
The Python application reads an Excel file and creates a dataframe. This data frame had corrupted data that is not usable to calculate values. For this reason I printed out the index location of these error, and removed the rows. Once the rows were deleted, I continued with the exercise to calculate the values. The application also creates an Excel file with the total accumulation of each column grouped by Company Name.
