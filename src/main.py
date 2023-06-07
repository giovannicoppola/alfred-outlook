"""
OUTLOOK-SUITE
A suite of tools for Outlook
Light rain, mist 🌦   🌡️+60°F (feels +60°F, 93%) 🌬️↙6mph 🌑&m Sat May 20 11:20:32 2023
W20Q2 – 140 ➡️ 224 – 9 ❇️ 356

"""

import sys
import os
import sqlite3
from consts import *
from util import log, timestampToText,  checkingTime, checkTimespan, checkJSON
import json
import subprocess



#initializing JSON output
result = {"items": [], "variables":{}}
MYSOURCE = os.getenv('mySource')

def fetchFolder ():
    """
    A function to fetch the folder name knowing its ID.
    Because going to the database using the function is too slow (and because the folders change rarely), decided to store the match in a JSON file in the data folder, and to update it every 30 days
    
    db = sqlite3.connect(OUTLOOK_DB_FILE)
    db.row_factory = sqlite3.Row
    rs = db.execute(f"SELECT Folder_Name from Folders WHERE Record_RecordID = {myFolderID}").fetchone()
    return rs[0]
    """

    with open(OUTLOOK_FOLDER_KEY_FILE, "r") as f:
        d = json.load(f)
    return d


def fetchAccounts ():
    
    with open(OUTLOOK_ACCOUNT_KEY_FILE, "r") as f:
        d = json.load(f)
    return d

def fetchContacts ():
    
    with open(OUTLOOK_CONTACTS_FILE, "r") as f:
        d = json.load(f)
    return d


def cleanSubject(mySubject):
    """
    a function to delete from subject used-defined strings which can take important real estate when results are returned.
    """
    for substring in WEED_TX:
            mySubject = mySubject.replace (substring, "")
    return mySubject

def showContacts(MY_INPUT, MY_CONTACTS,myQuery):
        MYOUTPUT = {"items": []}
        mySubset = [i for i in MY_CONTACTS if MY_INPUT.casefold() in i.casefold()]
        
        # adding a complete tag if the user selects it from the list
        if mySubset:
            for thisContact in mySubset:
                
                MYOUTPUT["items"].append({
                "title": f"{thisContact}",
                "subtitle": MY_INPUT,
                "arg": "",
                "variables" : {
                "selectedContact": thisContact,
                "myIter": True,
                "myQuery": myQuery,
                "mySource": 'contacts'
                    },
                "icon": {
                        "path": f"icons/contact.png"
                    }
                })
        else:
            MYOUTPUT["items"].append({
            "title": "no contacts matching",
            "subtitle": "try another query?",
            "variables" : {
                    
                    "myArg": MY_INPUT+" "
                    },
            "arg": "",
            "icon": {
                    "path": f"icons/Warning.png"
                }
            })
        print (json.dumps(MYOUTPUT))
        exit()

def compileSQL(myQuery,myFolderKeys,myAccountKeys,myContacts):
    """
    a function to parse the user's input and generate an SQL query that can be used to query the database
    """
    
    myElements = myQuery.split()
    # if len(myElements) > 1:
    conditions = []
    DEFAULT_SORT = 'DESC'
    
    for myElement in myElements:
        if myElement.startswith("from:"): #user is searching by sender
            #showContacts (myElement.split(":")[1],myContacts,myQuery)
            if myElement == 'from:me': #use the user-defined name string
                myString = MYSELF
            
            else: 
                myString = myElement.split(":")[1].strip()
                myString = myString.replace("_"," ")
                #myString = os.getenv('SelectedContact')
            conditions.append (f"Message_SenderList LIKE '%{myString}%'")
            
        elif myElement.startswith("to:"): #user is searching by sender
            
            if myElement == 'to:me': #use the user-defined name string
                myString = MYSELF
            
            else: 
                myString = myElement.split(":")[1].strip()
                myString = myString.replace("_"," ")
            conditions.append (f"Message_RecipientList LIKE '%{myString}%'")

        elif myElement.startswith("cc:"): #user is searching by sender
            
            if myElement == 'cc:me': #use the user-defined name string
                myString = MYSELF
            
            else: 
                myString = myElement.split(":")[1].strip()
                myString = myString.replace("_"," ")
            conditions.append (f"Message_CCRecipientAddressList LIKE '%{myString}%'")

        elif myElement == ("has:attach"): #user is searching for messages with attachments
                 
           
           conditions.append (f"Message_HasAttachment = 1")
        
        elif myElement.startswith("subject:"): #user is searching by subject
            
           
           myString = myElement.split(":")[1].strip()
           conditions.append (f"Message_NormalizedSubject LIKE '%{myString}%'")
        
        elif myElement.startswith("since:"): #user wants emails from the past xxx days
            
           myString = myElement.split(":")[1].strip()
           timeSpan = checkTimespan (myString)
           conditions.append (f"Message_TimeReceived > {timeSpan}")
        

        elif myElement.startswith("is:"): #read or unread
            
            myString = myElement.split(":")[1].strip()
           
            if myString == 'read':
                conditions.append (f"Message_ReadFlag = 1")
            
            elif myString == 'unread':
                conditions.append (f"Message_ReadFlag = 0")
            
            elif myString == 'important':
                conditions.append (f"Record_Priority = 1")
            elif myString == 'unimportant':
                conditions.append (f"Record_Priority = 5")

        elif myElement.startswith("is:"): #read or unread
            
            myString = myElement.split(":")[1].strip()
           
            if myString == 'read':
                conditions.append (f"Message_ReadFlag = 1")
            
            elif myString == 'unread':
                conditions.append (f"Message_ReadFlag = 0")

        elif myElement.startswith("folder:"): #user is searching by subject
           
           myString = myElement.split(":")[1].strip()
           myString = myString.replace("_"," ")
           myFolderID = next((key for key, value in myFolderKeys.items() if value.casefold() == myString.casefold()), None)
           
           if myFolderID:
            
            conditions.append (f"Record_FolderID = {myFolderID}")
        
        elif myElement.startswith("account:"): #user is searching by exchange account
           
           myString = myElement.split(":")[1].strip()
           myString = myString.replace("_"," ")
           myAccountID = next((key for key, value in myAccountKeys.items() if myString.casefold() in value.casefold()), None)
           
           if myAccountID:
            
            conditions.append (f"Record_AccountUID = {myAccountID}")
        
        elif myElement == ("--a"): #user wants to sort by decreasing date
           DEFAULT_SORT = 'ASC'
        
        elif myElement.startswith("-"): #user wants to exclude a word from preview or subject
            
           
           myString = myElement[1:].strip()
           conditions.append(f"(Message_Preview NOT LIKE '%{myString}%' AND Message_NormalizedSubject NOT LIKE '%{myString}%')")
        


        else: #default search: subject and preview (might want to customize that in workflow configuration)
            conditions.append(f"(Message_Preview LIKE '%{myElement}%' OR Message_NormalizedSubject LIKE '%{myElement}%')")
        
    
    if EXCL_FOLDERS:
        for myExclFolder in EXCL_FOLDERS:
            #try: 
            myFolderID = next((key for key, value in myFolderKeys.items() if value == myExclFolder.strip()), None)
                
            #except:
            #    myFolderKeys = fetchFolder ()
            #   myFolderID = next((key for key, value in myFolderKeys.items() if value == myExclFolder.strip()), None)
  
            if myFolderID:
                conditions.append(f"Record_FolderID <> {myFolderID}")
    
    
    # joining all the conditions
    conditions_str = " AND ".join(conditions)
    conditions_str = f" WHERE {conditions_str}"
    
    SQL_SORT = f'ORDER BY Message_TimeSent {DEFAULT_SORT}'
    sql = f"SELECT * FROM Mail {conditions_str} {SQL_SORT}"
    log (sql)
        
    return sql
    
    
def main():
    checkingTime()
    try: 
        myFolderKeys
    except:
        myFolderKeys = fetchFolder ()

    try: 
        myAccountKeys
    except:
        myAccountKeys = fetchAccounts ()

    try: 
        myContacts
    except:
        myContacts = fetchContacts ()

    # Check if the file has been updated today
    if checkJSON(OUTLOOK_SNOOZER_FILE):
        log("The JSON file has been updated today.")
    else:
        log("The JSON file has not been updated today. Updating....")
        subprocess.run(["python3", "unSnoozer.py"])


    myQuery = sys.argv[1]

    if MYSOURCE == "thread":
        mySQL = f"SELECT * FROM Mail WHERE Message_ThreadTopic = '{myQuery}' ORDER BY Message_TimeSent ASC"
    # elif MYSOURCE == "contacts":
    #     myQuery  = os.getenv('myQuery')
    elif myQuery:
        mySQL = compileSQL (myQuery,myFolderKeys, myAccountKeys,myContacts)
    else:
        mySQL = f"SELECT * FROM Mail ORDER BY Message_TimeSent DESC"
    
    handle(mySQL)




def handle(mySQL):
    

    db = sqlite3.connect(OUTLOOK_DB_FILE)
    db.row_factory = sqlite3.Row
    
    rs = db.execute(mySQL).fetchall()
    totCount = len(rs)
    
    myCounter = 0
    attIcon = ''
    for r in rs:
        myCounter += 1
        timeRec = timestampToText (r['Message_TimeSent'],"%Y-%m-%d %H:%M")
        timeRecF = timestampToText (r['Message_TimeSent'],"%A, %B %d, %Y %I:%M %p")
        if r['Message_ReadFlag'] == 0:
            readIcon = '⭐'
        else:
            readIcon = ''

        if r['Record_Priority'] == 1:
            urgentIcon = '‼️'
        elif r['Record_Priority'] == 5:
            urgentIcon = '🔽'
        
        else:
        
            urgentIcon = ''

        if r['Message_HasAttachment'] == 1:
            attIcon = '📎'
        if r['Message_NormalizedSubject']:
            MySubjectClean = cleanSubject(r['Message_NormalizedSubject'] )
        
        myPreview = f"{timeRecF}\nFrom: {r['Message_SenderList']}\nTo: {r['Message_RecipientList']}\n\nSubject: {MySubjectClean}\n\n{r['Message_Preview']}"
        myMessagePath = f"{OUTLOOK_MSG_FOLDER}{r['PathToDataFile']}"

        try: 
            myFolder = myFolderKeys [str(r['Record_FolderID'])]
        except:
            myFolderKeys = fetchFolder ()
            myFolder = myFolderKeys [str(r['Record_FolderID'])]
        if myFolder == "Snoozed":
            SnoozeString = "💤"
        else:
            SnoozeString = ""
        result["items"].append({
            "title": f"{readIcon}{MySubjectClean}{urgentIcon}",
            
            'subtitle': f"{myCounter}/{totCount:,} {attIcon} [{myFolder}] From: {r['Message_SenderList']} {timeRec} {SnoozeString}",
            'valid': True,
            "quicklookurl": '',
            'variables': {
                
            },
                "mods": {

                "control": {
                    "valid": 'true',
                    "subtitle": f"🧵 filter entire thread",
                    "arg": r['Message_ThreadTopic'],
                    'variables': {
                        "mySource": 'thread'   
                    }
                },
                "shift": {
                    "valid": 'true',
                    "subtitle": f"👀 show preview in large font",
                    "arg": myPreview
                }},
            "icon": {
                "path": f""
            },
            'arg': myMessagePath
                }) 
        
    if not rs:
        result["items"].append({
            "title": "No matches in your library",
            "subtitle": "Try a different query",
            "arg": "",
            "icon": {
                "path": "icons/Warning.png"
                }
            
                })
        
    print (json.dumps(result))



if __name__ == '__main__':
    main()
