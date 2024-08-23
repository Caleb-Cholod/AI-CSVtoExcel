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
import time
from functools import partial


#=====INTRO===========
#Written in python by Caleb C
#Program to take multiple csv files from the integritest system
#and generate excel files while deleting failiure tests.
#The program moves lot number excel files and used csv files into the shared strainrite sky folder
#Dependencies include tkinter, openpyxl, and shutil


#=====FILE NAME CONSTANTS=====

WB_COPY = 'master(Copy).xlsx'
FILES = 'CSVFiles'
USED_FILES = 'UsedCSVFiles'
LOTS = 'LotNumber'
BACKUPS = 'Backups'

SKY_LOCATION = r'\\sky\strainrite\AIT Reports'

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
    destination = FILES
    files = listdir(USED_FILES)
    for file in files:
        
        files_ = listdir(USED_FILES+"/"+file)
        for i in range(len(files_)):
            location = USED_FILES+"/"+file+"/"+files_[i]
            shutil.copy(location, destination)

    #Purge all files in oldCSV
    for file in files:
        shutil.rmtree(USED_FILES+"/"+file)

    #Change label
    num = listdir(FILES)
    change_label_text(len(num))

def copytree(src, dst, symlinks=False, ignore=None):
    for item in os.listdir(src):
        s = os.path.join(src, item)
        d = os.path.join(dst, item)
        if os.path.isdir(s):
            shutil.copytree(s, d, symlinks, ignore)
        else:
            print(s, d)

    

#=====ERROR/WARNING MESSAGES=====
def printError(index):
    if(index == 0):
        print("ERROR 01: Same Lot number and same Serial number found")

def printWarning(index):
    if(index == 0):
        print("WARNING 01: Empty Line Found in CSV")
    elif(index == 1):
        print("WARNING 02: Duplicate lot numbers in data")
    elif(index == 2):
        print("WARNING 03: No CSVs in CSVFiles, check AI_Excel>CSVFiles")
        
    

#=====MAIN============
def main():
    #Find CSV file(s)
    files = listdir(FILES)
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
        print("Opening CSV File #", index+1, "...")
        file = open(FILES + "/" + files[index], "r")
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
            os.rename(FILES+"/"+files[index], USED_FILES+"/"+CSVfileName+'/'+files[index])

        else:
            pass
            #Create a new CSV Folder
            os.mkdir(USED_FILES+"/"+CSVfileName)
            #Add CSV into our CSV folder
            os.rename(FILES+"/"+files[index], USED_FILES+"/"+CSVfileName+'/'+files[index])
            

    #Copy Workbook to a working copy
    #Save a copy of each excel spreadsheet
    csvs = listdir(LOTS)
    for csv in (csvs):
        location = LOTS + "/" + csv
        destination = BACKUPS + "/" + csv
        shutil.copyfile(location, destination)

    
    #Overwrite files in the strainrite shared folder location 
    for csv in (csvs):
        location = LOTS + "/" + csv
        destination = SKY_LOCATION + '\\LotNumber\\' + csv 
        shutil.copy(location, destination)


    #Overwrite CSV files in the strainrite shared folder location
    files = listdir(USED_FILES)
    for csv in (files):
        location = USED_FILES + "/" + csv
        destination = SKY_LOCATION + '\\CSVFiles\\' + csv 
        shutil.copytree(location, destination, dirs_exist_ok=True)

    #Save backup into strainrite shared folder
        


    #Change csv count to be 0
    change_label_text(0)

    #Finished
    print("Done")



    
#=====Tkinter Window=======
root = Tk()
root.title("Integrity CSV to Excel Tool")
root.geometry('375x400')
root.resizable(False, False)
root.configure(background='lightgray')
frm = ttk.Frame(root, padding=60)
frm.grid()


#Fonts
myfont = Font(family="Verdana", size=10)

#Labels
numfiles = str(len(listdir(FILES))) + " Files were found in AI_Excel>CSVFiles"
label_ = ttk.Label(frm, text=numfiles, font=myfont)
label_.grid(column=0, row=3, ipadx=0, ipady=5)

#TKinter Helper Functions
def change_label_text(num):
    label_.config(text=str(num)+" Files were found in AI_Excel>CSVFiles")


def refresh():
    num = listdir(FILES)
    change_label_text(len(num))

    

refreshImg = tk.PhotoImage(file='./refresh_.png')
smaller = refreshImg.subsample(4, 4)
refresh = ttk.Button(frm, command=refresh, image=smaller)
refresh.grid(column = 0, row=4, padx=0)


#TKinter Buttons
ttk.Button(frm, text="Combine files into spreadsheet", command=main).grid(column=0, row=2, padx=5, pady=20, columnspan=5, ipadx=15, ipady=15)
ttk.Button(frm, text="Quit", command=root.destroy).grid(column=0, row=12)
#ttk.Button(frm, text="Move Old CSV to new CSV folder", command=oldToNew).grid(column=0, row=6, pady=0, ipadx=0, ipady=0)
#ttk.Button(frm, text="Delete Lot Number Files", command=deleteLots).grid(column=0, row=7, ipadx=0, ipady=0)
#ttk.Label(frm, text="==========Debugging Tools===========").grid(column=0, row=5, ipadx=15, ipady=20)

#TKinter Image
photo = tk.PhotoImage(file='./image.png')
smaller_image = photo.subsample(2, 2)
image = ttk.Button(frm, image=smaller_image).grid(column=0, row=10, padx=0, pady=30)


root.mainloop()



