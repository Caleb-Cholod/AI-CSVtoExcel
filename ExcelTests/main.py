from openpyxl import load_workbook
from openpyxl import workbook
import shutil
from os import listdir
import os
import tkinter as tk
from tkinter import ttk
from tkinter import *
from tkinter.font import Font


#=====CLASSES=========


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
#============================================================

#=====HELPER FUNCTIONS=====

#Delete data in master
def resetMaster():
    print("Resetting master...")
    template = "master(Copy).xlsx"
    destination = "Spreadsheet.xlsx"
    shutil.copyfile(template, destination)

#Delete lotNo. excel files
def deleteLots():
    print("Deleting Files...")
    lots = listdir("LotNumber")
    for index in range(len(lots)):
        os.remove("LotNumber/" + lots[index])

#Move all in oldCSV to newCSV
def oldToNew():
    print("Moving Old CSV Files...")
    files = listdir("UsedCSVFiles")
    for index in range(len(files)):
        origin = "UsedCSVFiles/" + files[index]
        destination = "CSVFiles/" + files[index]
        os.rename(origin, destination)

#=====ERROR/WARNING MESSAGES=====
def printError(index):
    if(index == 0):
        print("ERROR 01, Same Lot # and same Serial #, deleting entry...")

def printWarning(index):
    if(index == 0):
        print("WARNING 01, Empty Line Found in CSV")
    elif(index == 1):
        print("Warning 02, Duplicate lot numbers in data")
        

#=====MAIN============
def main():
    print("Starting...")
    #Generate new excel file name based on csv and date


    #Find CSV file(s)
    files = listdir("CSVFiles")

    #open our master spreadsheet
    wb = load_workbook(filename = 'Spreadsheet.xlsx')

    #Get the next empty row
    maxRowCount = 1
    foundLimit = False

    while not foundLimit:
        
        if(wb['Sheet1']['A' + str(maxRowCount)].value == None):
            foundLimit = True
        else:
            maxRowCount += 1

    lineIncrement = -1


    #Open csv files and write to sheet
    #If a whole line is commas, end our read
    firstLine = True
    
    foundLotLimit = False
    activeWb = load_workbook(filename = 'Spreadsheet.xlsx')
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
                Lotfiles = listdir("LotNumber")

                if lot_no + ".xlsx" in Lotfiles:
                    
                    #Keep our active workbook open
                    #Find lotno. file, append to it
                    activeWb = load_workbook(filename = "LotNumber/" + lot_no + ".xlsx")
                    #if we are reopening this file, we need to find our current line
                    
                    while not foundLotLimit:
        
                        if(activeWb['Sheet1']['A' + str(lotRowCount)].value == None):
                            foundLotLimit = True
                        else:
                            lotRowCount += 1
                    
                else:
                    #create lotno. file
                    template = "master(Copy).xlsx"
                    destination = "LotNumber/" + lot_no + ".xlsx"
                    foundLotLimit = False
                    shutil.copyfile(template, destination)
                    
                    #Open new workbook as active workbook
                    activeWb = load_workbook(filename = "LotNumber/" + lot_no + ".xlsx")
                    
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
                        wb['Sheet1'][(chr(65 + i - offset)) + str(maxRowCount + lineIncrement)] = entry[i]
                        #Add entry to lot file
                        
                        activeWb['Sheet1'][(chr(65 + i - offset)) + str(lotRowCount)] = entry[i]

                        #Save workbook
                        activeWb.save("LotNumber/" + lot_no + ".xlsx")
                    else:
                        offset+=1

               #Change master ID to be correct
                wb['Sheet1']['A' + str(maxRowCount + lineIncrement)] = maxRowCount + lineIncrement - 1

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
        print(sorted((list(set(uniqueCSVS)))))
        #Add current csv to appropriate folder
        listdir("UsedCSVFiles")
        #close the file
        file.close()
        
        #move csv file into usedCSV folder
        #generate new csvfilename based on csv and date?
        os.rename("CSVFiles/"+files[index], "UsedCSVFiles/"+files[index])

    #Copy Workbook to a working copy

    #Close Workbook
    wb.save('Spreadsheet.xlsx')
    #Save a copy
    template = "Spreadsheet.xlsx"
    destination = "Backups/master.xlsx"
    shutil.copyfile(template, destination)


#Tkinter Setup
root = Tk()
root.title("Integrity CSV to Excel Tool")
root.geometry('400x500')
root.resizable(False, False)
root.iconbitmap('./image.png')
root.configure(background='lightgray')
frm = ttk.Frame(root, padding=0)
frm.grid()

#Fonts
myfont = Font(family="Verdana", size=10)

#Buttons Labels
ttk.Button(frm, text="Combine files into spreadsheet", command=main).grid(column=0, row=2, padx=5, pady=30,columnspan=5, ipadx=15, ipady=15)
numfiles = str(len(listdir("CSVFiles"))) + " Files were found."
ttk.Label(frm, text=numfiles, font=myfont).grid(column=0, row=3, ipadx=5, ipady=5)
ttk.Button(frm, text="Quit", command=root.destroy).grid(column=0, row=12)
ttk.Button(frm, text="Move Old CSV to new CSV folder", command=oldToNew).grid(column=0, row=5, pady=0, ipadx=0, ipady=0)
ttk.Button(frm, text="Delete master.xlsx Data", command=resetMaster).grid(column=0, row=6, ipadx=0, ipady=0)
ttk.Button(frm, text="Delete Lot Number Files", command=deleteLots).grid(column=0, row=7, ipadx=0, ipady=0)
ttk.Label(frm, text="==========Debugging Tools===========").grid(column=0, row=4, ipadx=5, ipady=10)

photo = tk.PhotoImage(file='./image.png')
image = ttk.Button(frm, image=photo).grid(column=0, row=10, padx=0, pady=30)

root.mainloop()


