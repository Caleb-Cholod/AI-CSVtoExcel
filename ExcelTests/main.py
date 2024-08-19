#=====IMPORTS===============
from openpyxl import load_workbook
from openpyxl import workbook
import shutil
from os import listdir
import os
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.font import Font

from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
from PyQt5.QtWidgets import (
    QApplication,
    QHBoxLayout,
    QLabel,
    QMainWindow,
    QPushButton,
    QStackedLayout,
    QVBoxLayout,
    QWidget,
    QCheckBox,
    QComboBox,
    QDoubleSpinBox,
    QLineEdit,
    QListWidget,
    QSlider,
    QSpinBox,
)
from PyQt5.QtGui import QPalette, QColor
from functools import partial


#=====INTRO===========
#Written in python by Caleb C
#Program to take multiple csv files from the integritest system
#and generate an excel file while deleting failiure tests.

#The program must be able to handle multiple CSV files put into
#the "NewCSV" folder and will move the added csv files into the
#"UsedCSV" folder

#The program will create a spreadsheet based on lot number, and should check
#that lot# doesnt exist
#The program will also add passed tests into a master sheet


#=====PYQT5 CLASSES==========
'''
class MainWindow(QMainWindow):
    def __init__(self):

        super(MainWindow, self).__init__()

        self.setWindowTitle("My App")

        pagelayout = QVBoxLayout()

        but1 = QPushButton("Press Me!")
        

        widget = QLabel("Hello")
        font = widget.font()
        font.setPointSize(30)
        widget.setFont(font)
        widget.setAlignment(Qt.AlignHCenter | Qt.AlignVCenter)

        self.setCentralWidget(widget)
        
    def activate_tab_1(self):
        self.stacklayout.setCurrentIndex(0)

    def activate_tab_2(self):
        self.stacklayout.setCurrentIndex(1)

    def activate_tab_3(self):
        self.stacklayout.setCurrentIndex(2)


class Color(QWidget):

    def __init__(self, color):
        super(Color, self).__init__()
        self.setAutoFillBackground(True)

        palette = self.palette()
        palette.setColor(QPalette.Window, QColor(color))
        self.setPalette(palette)
'''
#=====FILE NAME CONSTANTS=====

WB_COPY = 'master(Copy).xlsx'
USED_FILES = 'UsedCSVFiles'
LOTS = 'LotNumber'


#x = open(r'\\sky\strainrite\AIT Reports\424374.xlsx', 'r')
#SKY_LOCATION = '\\sky\strainrite\AIT Reports'

#=====HELPER FUNCTIONS=====

#Delete data in master
def resetMaster():
    print("Resetting master...")
    template = WB_COPY


#Delete lotNo. excel files
def deleteLots():
    print("Deleting Files...")
    lots = listdir(LOTS)
    for index in range(len(lots)):
        os.remove(LOTS+"/" + lots[index])

#Move all in oldCSV to newCSV
def oldToNew():
    #Copy files
    print("Moving Old CSV Files...")
    destination = "CSVFiles"
    files = listdir(USED_FILES)
    for file in files:
        files_ = listdir(USED_FILES+"/"+file)
        for i in range(len(files_)):
            location = USED_FILES+"/"+file+"/"+files_[i]
            shutil.copy(location, destination)

    #Purge all files in oldCSV
    for file in files:
        shutil.rmtree(USED_FILES+"/"+file)

    



#=====ERROR/WARNING MESSAGES=====
def printError(index):
    if(index == 0):
        print("ERROR 01: Same Lot number and same Serial number")

def printWarning(index):
    if(index == 0):
        print("WARNING 01: Empty Line Found in CSV")
    elif(index == 1):
        print("WARNING 02: Duplicate lot numbers in data")
    elif(index == 2):
        print("WARNING 03: No CSVs in CSVFiles, check folder")
        

#=====MAIN============
def main():
    #Find CSV file(s)
    files = listdir("CSVFiles")
    if(len(files) == 0):
        printWarning(2)

    print("Starting...")

    #Get the next empty row
    maxRowCount = 1
    foundLimit = False

    lineIncrement = -1


    #Open csv files and write to sheet
    #If a whole line is commas, end our read
    firstLine = True
    
    foundLotLimit = False
    activeWb = load_workbook(filename = WB_COPY)
    lotRowCount = 2
    
    for index in range(len(files)):
        print("Opening CSV File #", index, "...")
        file = open("CSVFiles/" + files[index], "r")
        #Lot File Variables
        uniqueCSVS = []
        for line in (file):
            if(not firstLine):
                entry = line.split(',')
                #check if line is all spaces
                if(entry.count("") >= 7):
                    printWarning(0)
                    break
                
                offset=0
                lotNoFlag = False

                #get lot_number
                lot_no = entry[11].rstrip()


                #Add lot number to list if unique
                uniqueCSVS.append(lot_no)
                #Check if lot number exists in lotfiles
                Lotfiles = listdir(LOTS)

                if lot_no + ".xlsx" in Lotfiles:
                    
                    #Keep our active workbook open
                    #Find lotno. file, append to it
                    activeWb = load_workbook(filename = LOTS+"/" + lot_no + ".xlsx")
                    #if we are reopening this file, we need to find our current line
                    
                    while not foundLotLimit:
        
                        if(activeWb['Sheet1']['A' + str(lotRowCount)].value == None):
                            foundLotLimit = True
                        else:
                            lotRowCount += 1
                    
                else:
                    #create lotno. file
                    template = WB_COPY
                    destination = LOTS + "/" + lot_no + ".xlsx"
                    foundLotLimit = False
                    shutil.copyfile(template, destination)
                    
                    #Open new workbook as active workbook
                    activeWb = load_workbook(filename = LOTS + "/" + lot_no + ".xlsx")
                    
                #before adding to our lot sheet, we want to check if lot number exists already in our activesheet
                for lot in range(lotRowCount):
                    if(activeWb['Sheet1']['H' + str(lot+1)].value == entry[9]):
                        printWarning(1)
                        if(activeWb['Sheet1']['J' + str(lot+1)].value == entry[11]):
                            lotNoFlag = True
                        
                if lotNoFlag:
                    printError(0)
                    break
                    
                for i in (range(len(entry))):
                    #print(entry[6])
                    #If we failed our diffusion teset rate in G, dont add this row.
                    if(float(entry[6]) <= 0.1):
                        print("    -failed diffusion test at entry", lineIncrement)
                        maxRowCount -= 1
                        lotRowCount -= 1
                        break
                    #If we failed our gross leak test in F, dont add this row.
                    if(float(entry[5]) <= 0.1):
                        print("    -failed gross leak test at entry", lineIncrement)
                        maxRowCount -= 1
                        lotRowCount -= 1
                        break
                        
                    #Add entry to spreadsheet (only entries A-G+J-L)
                    if(i < 7 or i > 8):
                        #Add entry to lot file
                        activeWb['Sheet1'][(chr(65 + i - offset)) + str(lotRowCount)] = entry[i]

                        #Save workbook
                        activeWb.save(LOTS + "/" + lot_no + ".xlsx")
                    else:
                        offset+=1

                lineIncrement += 1
                foundLotLimit = False
                lotRowCount = 2
                

            else:
                firstLine = False
                lineIncrement += 1
                
        #add number of lines in file to our maxRowCount
        
        maxRowCount += lineIncrement
        
        lineIncrement = -1
        firstLine = True

        #Get rid of duplicates and re-order out list of csvs
        uniqueCSVS = sorted(list(set(uniqueCSVS)))
        #Add current csv to appropriate folder
        CSVfileName = ''
        for name in (uniqueCSVS):
            CSVfileName += name
            if name != uniqueCSVS[-1]:
                CSVfileName += '-'

        #close the file
        file.close()
        
        if(os.path.exists(USED_FILES+"/"+CSVfileName)):
            #Add CSV into our CSV Folder
            os.rename("CSVFiles/"+files[index], USED_FILES+"/"+CSVfileName+'/'+files[index])

        else:
            pass
            #Create a new CSV Folder
            os.mkdir(USED_FILES+"/"+CSVfileName)
            #Add CSV into our CSV folder
            os.rename("CSVFiles/"+files[index], USED_FILES+"/"+CSVfileName+'/'+files[index])
            

    #Copy Workbook to a working copy
    #Save a copy of each excel spreadsheet
    csvs = listdir(LOTS)
    for csv in (csvs):
        location = LOTS + "/" + csv
        destination = "Backups/" + csv
        shutil.copyfile(location, destination)

#=====PYQT5================
'''
app = QApplication(sys.argv)

window = MainWindow()
window.show()

app.exec()

'''
#=====Tkinter Window=======
root = Tk()
root.title("Integrity CSV to Excel Tool")
root.geometry('400x500')
root.resizable(False, False)
#root.iconbitmap('./image.png')
root.configure(background='lightgray')
frm = ttk.Frame(root, padding=60)
frm.grid()

numfiles_ = str(len(listdir("CSVFiles"))) + " Files were found."

#Refresh Number of Files Text
def refresh():
    print("refreshing")
    global numfiles_
    label_.config(text=numfiles_)

    
#Fonts
myfont = Font(family="Verdana", size=10)

refreshImg = tk.PhotoImage(file='./refresh_.png')
smaller = refreshImg.subsample(3, 3)
#refresh = ttk.Button(frm, command=refresh, image=smaller).grid(column = 0, row=4, padx=0)


#Buttons Labels
ttk.Button(frm, text="Combine files into spreadsheet", command=main).grid(column=0, row=2, padx=5, pady=20, columnspan=5, ipadx=15, ipady=15)
numfiles = str(len(listdir("CSVFiles"))) + " Files were found."
label_ = ttk.Label(frm, text=numfiles, font=myfont).grid(column=0, row=3, ipadx=5, ipady=5)


ttk.Button(frm, text="Quit", command=root.destroy).grid(column=0, row=12)
ttk.Button(frm, text="Move Old CSV to new CSV folder", command=oldToNew).grid(column=0, row=6, pady=0, ipadx=0, ipady=0)
#ttk.Button(frm, text="Delete master.xlsx Data", command=resetMaster).grid(column=0, row=6, ipadx=0, ipady=0)
ttk.Button(frm, text="Delete Lot Number Files", command=deleteLots).grid(column=0, row=7, ipadx=0, ipady=0)
ttk.Label(frm, text="==========Debugging Tools===========").grid(column=0, row=5, ipadx=15, ipady=20)

photo = tk.PhotoImage(file='./image.png')
smaller_image = photo.subsample(2, 2)
image = ttk.Button(frm, image=smaller_image).grid(column=0, row=10, padx=0, pady=30)


root.mainloop()
    


