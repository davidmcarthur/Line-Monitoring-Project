#!/usr/bin/env python
# coding: utf-8

# In[20]:


#### Python Script developed by David McArthur 1/25/2021
#### For support please email david.mcarthur72@gmail.com
#### or david.mcarthur@onsemi.com

#### Don't mess with the functions/methods, any changed needed should occur in the Main
#### if the program isn't working please try to run the install package first.

################################################################################

###########  PROGRAM SETUP  ##########

#### Reference Libraries ####

import xlwings as xw   # XLWings enables Python integration with MS Excel
import os, re, os.path 
from pathlib import Path
from selenium import webdriver    # Selenium is used to remote control Firefox browsers
from selenium.webdriver import Firefox
from selenium.webdriver.firefox.options import Options
from datetime import datetime, timedelta, date
import webbrowser
import openpyxl

###############################################################################

##########  PROGRAM FUNCTIONS  ##########
 
#### Run Browser ####
# path is the folder where the geckodriver.exe file resides
# ftlink is the url for the fabtime chart your trying to run
# browser should close automatically on completion of method
#### Browser Setup ####
# folder is the desired download forlder for your particular chart.
# this method should be ran twice, once for EOH, once for WIP
def runbrowser(path, ftlink, folder):
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", folder)
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    #profile.set_preference("browser.download.manager.closeWhenDone", True)
    driver = webdriver.Firefox(executable_path=path, firefox_profile=profile) 
    driver.get(ftlink)
    #driver.teardown() 

### Create a file folder ####
# The intent is to create two folders, one for each chart we're working with.
# foldername is the location of the chart you're looking to utilize
def createfolder(folder):
    p = Path(folder)
    try:
        p.mkdir()
    except FileExistsError as exc:
        deletefiles(folder)
        deletefolder(folder)
        p = Path(folder)
        print(exc)

#### Find latest file #####
# Cleaning up the files and deleting the folders is critical
# we really only want the one file we're working with in existence
# at one time. All files and folders should be destroyed once the 
# data has been transfered
# foldername is the location of the chart you're looking to utilize
def findfile(folder):
    entries = os.listdir(folder)
    print(entries)
    return entries[0]

#### Delete all files in path ####
# File cleanup is critical to this program to eliminate recycling old data.
# Perform file and folder deletion upon completion of data extraction.
def deletefiles(folder):
    try:
        for root, dirs, files in os.walk(folder):
            for file in files:
                os.remove(os.path.join(root, file))
    except OSError as e:
        print(f'Error: {trash_dir} : {e.strerror}')

#### Delete file folder ####
# Must be ran once the files in the folder have been deleted
# last step of the Cleanup
def deletefolder(folder):
    try:
        os.rmdir(folder)
    except OSError as e:
        print(f'Error: {trash_dir} : {e.strerror}')

################################################################################

###########  MAIN  ##########

def main():
    #### Paths ####

    # needs to point to the location of the geckodriver.exe, this file is where the program interface to Firefox resides.
    path = "C:\\users\\zbmv6f\\Desktop\\LineMonitoring\\geckodriver.exe"

    lm_excel_name = "C:\\Users\\zbmv6f\\Desktop\\LineMonitoring\\LineMonitoring.xlsx"

    # needs to point to the firefox link to the correct EOh chart
    ft_chart_eoh = "http://fabtimeusuprd.onsemi.com/FabTime717/GetTable.asp?AreasLike=C-FAB10%2C%20F-POSTFAB&TimeUMID=2&FactoryID=1&CustomTitle=EOH&EditChart=0&ReworkChoice=1&QueueChoice=1&StripeColor=Black&ChartSortField3=ObjectPlusDescription&ChartSortField4=CrossObjectPlusDescription&Grid=None&LogHomePageTab=.0001%20Line%20Report&CrossObjectTypeID=1002&StripeAxis=Y&DataValues=None&Horizontal=0&ActiveMode=Javascript&DataRows=250&ByObjectTypeID=16&LogChart=WIPStkPareto&OwnersLike=~TE*%2C~MA*%2C~NR*&Chart=WIPStkPareto&AgeChoice=2&Refresh=1440&HoldChoice=1&PageMode=2&Width=1326.25&Height=897.75&GoalLogin=FTUsuAdm&GoalUserID=149&SkipBuildChart=1&HideTableCellLinks=1&HideColumnControls=1&RowsBetweenHeadings=0&ContentType=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet&AllTableData=1&WriteXLS=1"

    # This link needs date/time handling. If any errors in program this is a suspect area
    ftlink_dty = date.today() + timedelta(days=-1)
    ftlink_dtt = date.today()
    # Fabtime chart for moves link .001 on WILIAM WARD
    ft_chart_moves = "http://fabtimeusuprd.onsemi.com/FabTime717/GetTable.asp?AreasLike=C-FAB10%2C%20F-POSTFAB&TimeUMID=2&FactoryID=1&CustomTitle=MOVES&EditChart=0&FiscalPeriodType=Day&ReworkChoice=1&StripeColor=Black&ChartSortField3=ObjectPlusDescription&ChartSortField4=IsPredicted&ChartSortField5=CrossObjectPlusDescription&Grid=None&LogHomePageTab=.0001%20Line%20Report&CrossObjectTypeID=1002&RangeRelativeFrom=-1&RangeRelativeTo=-1&EndTime={0}%207%3A0%3A0&StripeAxis=Y&DataValues=None&Horizontal=0&ActiveMode=Javascript&DataRows=250&ByObjectTypeID=16&LogChart=MovesStkPareto&BoundaryHHMM=07%3A00&OwnersLike=~TE*%2C~MA*%2C~NR*&Chart=MovesStkPareto&Refresh=1440&DateChoiceMethod=Relative&StartTime={1}%207%3A0%3A0&PageMode=2&Width=1326.25&Height=897.75&GoalLogin=FTUsuAdm&GoalUserID=149&SkipBuildChart=1&HideTableCellLinks=1&HideColumnControls=1&RowsBetweenHeadings=0&ContentType=application%2Fvnd.openxmlformats-officedocument.spreadsheetml.sheet&AllTableData=1&WriteXLS=1".format(ftlink_dtt, ftlink_dty)

    #### Folder Setup ####
    # This folder will keep the temp excel file for eoh
    fldr_eoh = "C:\\users\\zbmv6f\\desktop\\LineMonitoring\\EOH"
    # This folder will keep the temp excel file for wip
    fldr_moves = "C:\\users\\zbmv6f\\desktop\\LineMonitoring\\MOVES" 

    ## PRE-CLEAN FILES AND FOLDERS
#    deletefiles(fldr_eoh)
#    deletefiles(fldr_moves)

#    deletefolder(fldr_eoh)
#    deletefolder(fldr_moves)

    ## CREATE FOLDERS
    createfolder(fldr_eoh)
    createfolder(fldr_moves)

    ## BROWSER SETUP & DOWNLOADS
    runbrowser(path, ft_chart_eoh, fldr_eoh)
    runbrowser(path, ft_chart_moves, fldr_moves)

    ### this takes a ton of time, need to wait here

    ## XLWINGS 

    ###### need to add functionality to check that file !=GetTable.
    file_moves = findfile(fldr_moves)
    file_eoh = findfile(fldr_eoh)

    wordmove = file_moves[0:7]
    wordeoh = file_eoh[0:7]

    while wordmove != 'ONSEMI_':
        if file_moves[0:7] == "GetTable":
            deletefiles(fldr_moves)
            runbrowser(path, ft_chart_moves, fldr_moves)
            file_moves = findfile(fldr_moves)
            wordmove = file_moves[0:7]

    while wordeoh != 'ONSEMI_':
        if file_eoh[0:7] == 'GetTable':
            deletefiles(fldr_eoh)
            runbrowser(path, ft_chart_eoh, fldr_eoh)
            file_eoh = findfile(fldr_eoh)
            wordeoh = file_eoh[0:7]
    wb_r2lr = xw.Book(lm_excel_name)

    # Point to main line report xlsx sheets to be manipulated
    lr_master = wb_r2lr.sheets['Master Tab']
    lr_eoh = wb_r2lr.sheets['EOH']
    lr_moves = wb_r2lr.sheets['Moves']

    wbmoves = xw.Book("{0}\\{1}".format(fldr_moves, file_moves))
    moves = wbmoves.sheets[0]
    wbeoh = xw.Book("{0}\\{1}".format(fldr_eoh, file_eoh))
    eoh = wbeoh.sheets[0]

    # could us a timer here

    # MOVE EOH TO BOH

    l1 = lr_master.range("H2").options(expand = "down").value
    lr_master.range("I2").options(transpose = True).value = l1
    ####

    eoh_from = ['E8', 'F8', 'H8', 'I8', 'J8', 'B8']
    eoh_to =   ['B2', 'C2', 'D2', 'E2', 'F2', 'G2']

    for f, t in zip(eoh_from, eoh_to):
        l_eoh = eoh.range('{0}'.format(f)).options(expand = "down").value
        lr_eoh.range('{0}'.format(t)).options(transpose = True).value = l_eoh

    moves_from = ['C10', 'F10', 'H10', 'J10', 'K10', 'L10', 'G10']
    moves_to =   ['H2',  'B2',  'D2',  'E2',  'F2',  'G2', 'C2']

    for f, t in zip(moves_from, moves_to):
        l_moves = moves.range('{0}'.format(f)).options(expand = "down").value
        lr_moves.range('{0}'.format(t)).options(transpose = True).value = l_moves

    # Need to close the downloaded workbooks
    wbmoves.close()
    wbeoh.close()
    wb_r2lr.save()

    ## CLEANUP FILES AND FOLDERS
    deletefiles(fldr_eoh)
    deletefiles(fldr_moves)

    deletefolder(fldr_eoh)
    deletefolder(fldr_moves)

if __name__ == "__main__":
    main()


# In[25]:




