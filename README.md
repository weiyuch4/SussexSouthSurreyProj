# Autoplan Renewal Letters Generator

The program is designed to read data from a database containing information about the policyholder, the vehicle being insured, the policy's expiration date and more. The program will then automatically generate renewal letters by replacing specific keywords in a pre-existing Word document template with all the relevant details about the policy, such as the policyholder's name, address, the vehicle make and model, and the policy's renewal date.

## Requirements

- Python (version 3.6 or higher)
- pyodbc module (version 4.0.32 or higher)
- python-docx module (version 0.8.11 or higher)
- Microsoft Word (version 2013 or higher)

## Installation

1. Clone or download the repository to your local machine.
2. Install Python (version 3.6 or higher) from the official website: 
https://www.python.org/downloads/
3. Install the pyodbc and python-docx modules by running the following command in your terminal or command prompt:
```
pip install pyodbc
```
```
pip install python-docx
```

## Usage

1. Run the `renewal.bat` file to execute the program.
2. When prompted, enter the start date, end date and your name for the letter generation. The dates should be in DD/MM/YYYY format.
3. The program will generate all autoplan renewal letters that are due withthin the specified date range by replacing the specific words in the template document with the customer data from the database.
4. The generated letters will be saved in a seperate directory named `TEMP`.

## Example

Suppose you want to generate all autoplan renewal letters that are due between March 1, 2023 and April 1, 2023. You would run the `renewal.bat` file. You will then be prompted to enter the start date and end date in the format of DD/MM/YYYY. You would enter the following:
```
Enter Start Date (DD/MM/YYYY): 01/03/2023
Enter End Date (DD/MM/YYYY): 01/04/2023
```
The program will return the number of renewals found within the specified date range and prompt you to enter your name:
```
Count: 24
Enter your full name: YourName
```
The program will then generate all the renewal letters that are due between March 1, 2023 and April 1, 2023.

