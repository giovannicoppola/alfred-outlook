from consts import *
import sys
from datetime import datetime, date, timedelta
import sqlite3
import json
import time
import re
import os


def log(s, *args):
    if args:
        s = s % args
    print(s, file=sys.stderr)



def checkTimespan (myTimeSpan):

    # current epoch number
    epoch = time.time()

    # convert to struct_time object
    st = time.localtime(epoch)

    
    pattern = r'^(\d+)([wm]?)$' #chekcing if the user used w or m
    match = re.match(pattern, myTimeSpan)

    if match:
        num_str, letter = match.groups()
        num = int(num_str)
        if letter == 'w':
            num *= 7
        elif letter == 'm':
            num *= 30
            
    else:
        num = myTimeSpan
    
    # subtract X days (for example, 10 days)
    dt = datetime.fromtimestamp(time.mktime(st)) - timedelta(days=num)

    # set hour, minute and second to zero
    dt = dt.replace(hour=0, minute=0, second=0)

    # convert back to epoch number
    new_epoch = time.mktime(dt.timetuple())

    
    
    
    # print the result
    return int(new_epoch)



def checkJSON(file_path): #check if the JSON file has been updated today
    
    if os.path.exists(file_path) and os.path.getsize(file_path) > 0:
        log ("checking the JSON snoozer file...")
        
        # Get the modification timestamp of the file
        timestamp = os.path.getmtime(file_path)

        # Convert the timestamp to a datetime object
        modification_date = datetime.fromtimestamp(timestamp)

        # Get today's date
        today = datetime.now().date()

        # Compare the modification date with today's date
        return modification_date.date() == today


def getFolderData():
    myFolders = {}
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    #db.row_factory = sqlite3.Row
    cur = db.cursor()
    cur.execute(f"SELECT Record_FolderID from Mail")
    folderIDs = set([row[0] for row in cur.fetchall()])
    for myFolder in folderIDs:
        rs = cur.execute(f"SELECT Folder_Name from Folders WHERE Record_RecordID = {myFolder}").fetchone()
        myFolders[myFolder] = rs[0]
    with open(OUTLOOK_FOLDER_KEY_FILE, "w") as f:
            json.dump(myFolders, f, indent=4)              
        
    log (myFolders)


def getAccountData():
    myAccounts = {}
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    
    cur = db.cursor()
    cur.execute(f"SELECT Record_AccountUID from Mail")
    accountIDs = set([row[0] for row in cur.fetchall()])
    for myAccount in accountIDs:
        #log (myAccount)
        rs = cur.execute(f"SELECT Account_Name from AccountsExchange WHERE Account_MailAccountUID = {myAccount}").fetchone()
        if rs:
            myAccounts[myAccount] = rs[0]
    with open(OUTLOOK_ACCOUNT_KEY_FILE, "w") as f:
            json.dump(myAccounts, f, indent=4)              
        
    log (myAccounts)



def getContactList():
    """
    compiling the autocomplete contact list from the list of senders currently in the database
    """
    myContacts = []
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    
    cur = db.cursor()
    rs = cur.execute(f"SELECT Message_SenderList from Mail")
    if rs:
        for myRs in rs:
            #log (myRs)
            try:
                currentContacts = myRs[0].split(";")
                for thisContact in currentContacts:
                    thisContact = thisContact.strip()
                    thisContact = thisContact.replace("'","")
                    if thisContact in myContacts:
                        continue
                    else:
                        myContacts.append(thisContact)
            except:
                continue
    myContacts.sort()
    with open(OUTLOOK_CONTACTS_LIST_FILE, "w") as f:
            json.dump(myContacts, f, indent=4)              
        
    #log (myContacts)

def getContactBook():
    """
    getting the autocomplete contact list from the address book
    """
    myContacts = []
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    
    cur = db.cursor()
    rs = cur.execute(f"SELECT Contact_DisplayName from Contacts")
    if rs:
        for myRs in rs:
            #log (myRs)
            try:
                currentContacts = myRs[0].split(";")
                for thisContact in currentContacts:
                    thisContact = thisContact.strip()
                    thisContact = thisContact.replace("'","")
                    if thisContact in myContacts:
                        continue
                    else:
                        myContacts.append(thisContact)
            except:
                continue
    myContacts.sort()
    with open(OUTLOOK_CONTACTS_BOOK_FILE, "w") as f:
            json.dump(myContacts, f, indent=4)              
        
    #log (myContacts)


def fetchRecordID (myID):
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    cur = db.cursor()
    rs = cur.execute(f"SELECT Message_MessageID FROM Mail WHERE Record_RecordID = {myID}").fetchone()
    return (rs[0])
    
def fetchEmailID (myID):
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    cur = db.cursor()
    rs = cur.execute(f"SELECT Record_RecordID FROM Mail WHERE Message_MessageID = '{myID}'").fetchone()
    return str(rs[0])


def checkingTime ():
## Checking if the database needs to be built or rebuilt
    timeToday = date.today()
    if not os.path.exists(OUTLOOK_FOLDER_KEY_FILE):
        log ("Database missing ... building ⏳")
        getFolderData()
        
    else: 
        databaseTime= (int(os.path.getmtime(OUTLOOK_FOLDER_KEY_FILE)))
        dt_obj = datetime.fromtimestamp(databaseTime).date()
        time_elapsed = (timeToday-dt_obj).days
        log (f"{time_elapsed} days from last update")
        if time_elapsed >= REFRESH_RATE:
            log ("rebuilding database ⏳...")
            getFolderData()
            log ("done 👍")


def timestampToText (epoch_timestamp,formatString):

    # Convert epoch timestamp to datetime object
    dt = datetime.fromtimestamp(epoch_timestamp)
    
    # Format the datetime object to the desired format
    formatted_date = dt.strftime(formatString)
    
    return formatted_date

