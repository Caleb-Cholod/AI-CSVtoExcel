Program written by Caleb Cholod, this program takes CSV files from the Integrity Testing system and sorts them into
an excel spreadsheet according to lot number, automatically deleting entries that fail the diffusion test. 

Dependencies:
Python 3.12 - https://www.python.org/downloads/

OpenPyXL - pip install openpyxl
pyqt5 - pip install pyqt5


Usage:
Place new CSV files into the NewCSV folder
Double click the main.py program
Number of CSV Files are shown, if the number is correct then click the combine files into spreadsheet button
Warnings/Errors are displayed in terminal


Warnings:
Make sure master sheet is titled Sheet1, and make sure files are properly closed on system
Make sure no excel spreadsheets are open
