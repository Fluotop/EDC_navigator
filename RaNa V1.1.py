"""""
;===================================================================================================
; Title: RaNa
; Author: Ben De Maesschalck
; Version: V1.1 (2/May/2022)
; Language: Python V3.7.9
;===================================================================================================
Dependencies to install through pip (or conda):
- pandas
- tkinter
- selenium
- pyautogui
- re
- bindglobal
- pynput
"""

"""
VERSION HISTORY
V1.0: Release
V1.1: Added try block in click folder - reduces errors for orange queries significantly.
"""

import sys
import time
from tkinter.filedialog import askopenfilename
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.wait import WebDriverWait
from selenium.common.exceptions import NoSuchElementException, TimeoutException, WebDriverException
import re
import pyautogui
from tkinter import *
from bindglobal import BindGlobal
import pynput

#########################################
# Display Settings and Error Suppression#
#########################################
sys.tracebacklimit = 10
pd.options.mode.chained_assignment = None
pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.max_colwidth', None)
pd.set_option('display.width', 1000)
pd.set_option('display.colheader_justify', 'center')
pd.set_option('display.precision', 3)

###############################################################
# Constants with default values. These are free to be changed##
###############################################################

VERSION = "RaNa v1.1"
WIDTH = pyautogui.size().width
HEIGHT = pyautogui.size().height
GEOMETRY = str(round(WIDTH * 2 / 3)) + "x" + str(round(HEIGHT * 0.13))
COUNT = -1

#####################################################################

"""
Creating the .xlsx file selection window.
"""
# Launch tk app and hide main window
main_window = Tk()
main_window.withdraw()

# Open fileselection .xlsx file from system

filename = askopenfilename(title="Columns Required: Subject Name, Folder Name, Site and Page",
                           filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))

# Settings for main window
main_window.geometry(GEOMETRY)
main_window.title(VERSION)
label = Label(main_window, text="Press Alt + q To Start")
text = Text(main_window, width=WIDTH, height=50)
label.pack()
text.pack_forget()

"""Converts the subject from a number to a string, adds a leading 0 if the site number starts with a zero 0 (leading 
zero's are not kept by excel) and inserts a column with name equal the the VERSION if the column is not present. The 
VERSION column is used to indicate resolved queries or to flag unresolved queries. This function is always and only 
called by the function load_data(). 

:parameter
df: pandas.core.frame.DataFrame
    Original dataframe as provided in the fileselection window.
    
:returns
df: pandas.core.frame.DataFrame
    df where column Subject Name which is converted from a int to a str. 
    Contains a column VERSION which is used to save progress.

"""


def prepare_data(df):
    global numberOfColumns
    global AMOUNT_OF_QUERIES

    df["Subject Name"] = df["Subject Name"].astype(str)
    df['Site'] = df['Site'].apply(str)
    sitenumbers = df['Site'].str.extract('(\d+)', expand=False).str.strip()


    for index, row in df.iterrows():
        if (df["Subject Name"][0][:1] >= "0" and df["Subject Name"][0][:1] <= "9"):
            if row["Subject Name"][:len(sitenumbers[index])] != sitenumbers[index]:
                df["Subject Name"][index] = "0" + row["Subject Name"]

    if VERSION in df.columns:
        AMOUNT_OF_QUERIES = df[VERSION].gt(0).idxmax()

    elif VERSION not in df.columns:
        df.insert(0, VERSION, 0)
        AMOUNT_OF_QUERIES = len(df)

    numberOfColumns = len(df.columns)
    return df


"""Reads the provided excel file looking for certain column combination. At least the columns: Subject Name, 
Folder Name and Page have to be present for a xlsx file to be considered valid. If there is a match a df will be 
prepared with function prepare_data(), the first unresolved query is determined via the VERSION column if present and 
columns to be visualized in the tk output window are selected. 

:parameter
filename: str
    The filepath of the selected .xlsx file.
"""


def load_data(filename):
    global df
    global first3
    global DFSUBSET

    df = pd.read_excel(filename)
    if {"Subject Name", "Folder Name", "Site", "Page", "Field", "Record Sequence", "Query Text"}.issubset(df.columns):
        df = prepare_data(df)
        df["Short Query Text"] = df["Query Text"].str[:17]
        displayedFields = ["Subject Name", "Folder Name", "Page", "Field", "Record Sequence", "Short Query Text"]

    elif {"Subject Name", "Folder Name", "Site", "Page", "Record Sequence"}.issubset(df.columns):
        df = prepare_data(df)
        displayedFields = ["Subject Name", "Folder Name", "Page", "Record Sequence"]
    
    elif {"Subject Name", "Folder Name", "Site", "Page"}.issubset(df.columns):
        print("meme")
        df = prepare_data(df)
        displayedFields = ["Subject Name", "Folder Name", "Page"]
    


    first3 = pd.DataFrame(columns=displayedFields)
    DFSUBSET = df[displayedFields]


###################################################################################################
# Enter function, load data and set the values for global variables AMOUNT_OF_QUERIES, first3 and DFSUBSET#
###################################################################################################
load_data(filename)
AMOUNT_OF_QUERIES
DFSUBSET
first3

##################################################################################################


# Show main window (tk)
main_window.deiconify()

# Open browser(medidata.com) and set size.
driver = webdriver.Edge(executable_path="edgedriver.exe")
driver.set_window_position(0, 0)
driver.set_window_size(WIDTH * 0.67, HEIGHT * 0.55)
driver.get("https://login.imedidata.com/login?service=https%3A%2F%2Fwww.imedidata.com%2F")

"""Returns to the main search screen for the MediData study, changes the settings to Search By Subject, enters the 
next Subject Name in the search box, presses enter and clicks on the first identical match listed. This function is 
only called by load_next_query(). 

:parameter
subjectname str:
    The subjectname corresponding to the query that needs to be loaded.
"""


def search_subject(subjectname):
    # Returns to the main search screen for the MediData study
    environments = driver.find_element(by=By.ID, value="study_environments_menu")
    environments.click()
    first_study_environment = driver.find_element(by=By.XPATH,
                                                  value="//*[@id='study_environments']/ul/li[2]/ul/li[2]/a")
    first_study_environment.click()

    # Changes the settings to Search By Subject
    search_By_Site = WebDriverWait(driver, 12).until(
        EC.element_to_be_clickable((By.XPATH, "//button[text()='Search By Site']")))
    search_By_Site.click()
    search_By_Subject = driver.find_element(by=By.XPATH,
                                            value="/html/body/div[6]/div/div/div[2]/div/div[1]/form/span[1]/ul/li[2]")
    search_By_Subject.click()

    # Enters Subject Name and Search
    search_field = driver.find_element(by=By.XPATH, value="/html/body/div[6]/div/div/div[2]/div/div[1]/form/input")
    search_field.send_keys(subjectname)
    search_field.send_keys(Keys.ENTER)

    # Click the result that is identical to the Subject Name
    xpathID = "//td[.=\"" + subjectname + "\"]"
    search_result = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, xpathID)))
    search_result.click()


"""Changes the subject via the dropdown menu (when on a patient page). Used for consecutive queries that have 
identical sites. This function is only called by load_next_query(). """


def change_subject(subjectname):
    # Open menu
    subjects = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.ID, "app_subjects_menu")))
    subjects.click()
    # Click next subject
    nextSubject = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, subjectname)))
    nextSubject.click()


"""Navigates the folders using the bottons on the left (when on a patient page). A distinction must be made for 
folders containing subfolders (Cycles) and for folders with names containing leading or trailing spaces. This 
function is only called by load_next_query(). """


def click_foldername(foldername):
    # Folders containing the name Cycle
    if "Cycle" in foldername:
        if foldername[:7] == "Cycle 1":
            mainfolder = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, "Cycle 1")))
            mainfolder.click()

        # Extract the cycle number when cycle != 1
        if foldername[:7] != "Cycle 1":
            cycle = re.search("Cycle \d+", foldername).group()
            mainfolder = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, cycle)))
            mainfolder.click()
        try:
            # Click the subfolder if present
            folder = WebDriverWait(driver, 3).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, foldername)))
            folder.click()
        except:
            return

    # Folders containing trailing spaces
    elif foldername[-1] == " ":
        folder = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, foldername[:-1])))
        folder.click()

    # Default
    else:
        folder = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.PARTIAL_LINK_TEXT, foldername)))
        folder.click()


"""
Navigates the pages(=sub[(subfolders)] using the buttons on the left (when on a patient page).
This function is only called by load_next_query().
"""


def click_pagename(pagename):
    page = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.LINK_TEXT, pagename)))
    page.click()


"""
Checks the differences between the previous and next query and determines the fastest way to navigate to the next query.
Default is navigating from root.
This function is only called by write_display().
"""


def load_next_query():
    nextSubject = df.iloc[COUNT]["Subject Name"]
    nextFolder = df.iloc[COUNT]["Folder Name"]
    nextPage = df.iloc[COUNT]["Page"]
    nextSite = re.search("\d+", df.iloc[COUNT]["Site"]).group()

    # First query has to nothing to be compared to, others get compared to the previous one
    if COUNT == 0:
        previousSubject = ""
        previousFolder = ""
        previousPage = ""
        previousSite = ""
    else:
        previousSubject = df.iloc[COUNT - 1]["Subject Name"]
        previousFolder = df.iloc[COUNT - 1]["Folder Name"]
        previousPage = df.iloc[COUNT - 1]["Page"]
        previousSite = re.search("\d+", df.iloc[COUNT - 1]["Site"]).group()

    # If the next page is the same as the current one don't navigate and display a green color
    if nextSite == previousSite and nextSubject == previousSubject and nextFolder == previousFolder and nextPage == previousPage:
        main_window.configure(bg="green")
        time.sleep(0.25)

    # Elif the next folder is the same as the current one only navigate to the different page and display a blue color
    elif nextSite == previousSite and nextSubject == previousSubject and nextFolder == previousFolder:
        main_window.configure(bg="sky blue")
        click_pagename(nextPage)

    # Elif the next subject is the same as the previous one only navigate to the new folder and new page and display
    # an orange color
    elif nextSite == previousSite and nextSubject == previousSubject:
        main_window.configure(bg="orange")
        print(previousFolder, nextFolder, nextPage)
        click_foldername(previousFolder)
        time.sleep(0.25)
        click_foldername(nextFolder)
        click_pagename(nextPage)

    # Elif the next site is the same as previous one only switch subjects and display a black color.
    elif nextSite == previousSite:
        main_window.configure(bg="black")
        change_subject(nextSubject)
        time.sleep(7)
        click_foldername(previousFolder)
        time.sleep(0.25)
        click_foldername(nextFolder)
        click_pagename(nextPage)

    # Else the next subject is not the same then start navigating from root
    else:
        main_window.configure(bg="white")
        search_subject(nextSubject)
        click_foldername(nextFolder)
        click_pagename(nextPage)


"""
Writes a 1 in the "VERSION" column of the excel file if alt + q was pressed and a 2 if alt + r was pressed.

:parameter
flag Boolean:
    Write a 1 in the first column of the original excel file or not.
previousQuery int:
    The df index of the previousQuery.
"""


def write_to_file(flag, previousQuery):
    try:
        if COUNT != 0:
            if not flag:
                df.at[previousQuery, VERSION] = 1
            elif flag:
                df.at[previousQuery, VERSION] = 2
            dfcopy = df.copy()
            dfcopy.sort_values(by=[VERSION, "Days Unresolved"], ascending=[True, False], inplace=True)
            dfcopy.iloc[:, :numberOfColumns].to_excel(filename, index=False)
    except (PermissionError,KeyError):
        return
        
"""
Will check if all queries have been loaded, if so pressing hotkeys will no longer have an effect and a purple text box will be displayed. If there are untouched queries available
then it will start the process of loading the next query and will display information about the loading process in the TK window. When query is not found will display a red color in the TK
Window and display info about the query. When query is found info about the three previous queries is displayed (if possible) in the TK window.

:parameter
flag Boolean:
    Write a 1 in the first column of the original excel file or not.
"""


def write_display(flag):
    global COUNT
    global first3
    COUNT += 1
    previousQuery = COUNT - 1
    nextQuery = COUNT

    if COUNT >= AMOUNT_OF_QUERIES:
        if COUNT == AMOUNT_OF_QUERIES:
            write_to_file(flag, previousQuery)
        label.config(text='No More Queries To Load')
        label.configure(bg="SlateBlue1")
        main_window.configure(bg="white")
        return

    write_to_file(flag, previousQuery)

    try:
        label.config(text='Navigating To Next Query...')
        label.configure(bg="salmon")
        main_window.configure(bg="white")
        print("Navigating To Next Query...")
        load_next_query()

    except (NoSuchElementException, TimeoutException, WebDriverException):
        main_window.configure(bg="red")
        label.config(
            text="Navigation failed please navigate to this query manually. Press alt + q to continue to the next query.")
        label.configure(bg="dark sea green")
        text.delete("1.0", END)
        text.insert("1.0", DFSUBSET.iloc[nextQuery, :])
        text.pack()
        print("Navigation failed please navigate to this query manually. Press alt to continue to the next query.")
        print(DFSUBSET.iloc[nextQuery, :])
        first3 = first3.append(DFSUBSET.iloc[nextQuery])

    else:
        if COUNT <= 2:
            first3 = first3.append(DFSUBSET.iloc[nextQuery])
            label.config(text="Query History")
            label.configure(bg="dark sea green")
            text.delete("1.0", END)
            text.insert("1.0", first3)
            text.pack()
            print(first3)
        elif COUNT > 2:
            label.config(text="Query History")
            label.configure(bg="dark sea green")
            text.delete("1.0", END)
            text.insert("1.0", DFSUBSET.iloc[previousQuery - 1: nextQuery + 1])
            text.pack()
            print(DFSUBSET.iloc[previousQuery - 1: nextQuery + 1])


"""
Set whether a 1 should be written in the first column of the original excel file.

:parameter
e bindglobal.PynputEvent:
    required as per bindglobal docs
"""


def no_mark(e):
    mark_query = False
    write_display(mark_query)


"""
Set whether a 2 should be written in the first column of the original excel file.

:parameter
e bindglobal.PynputEvent:
    required as per bindglobal docs
"""


def mark(e):
    mark_query = True
    write_display(mark_query)


# Hotkey setup
bg = BindGlobal()
bg.start()
bg.gbind("<Alt_L-KeyRelease-q>", no_mark)
bg.gbind("<Alt_L-KeyRelease-r>", mark)
main_window.mainloop()