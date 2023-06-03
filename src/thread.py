"""
SHOW THREAD
"""

import sys
import os
import sqlite3
from consts import *
from util import log, timestampToText
import json
from config import MYSELF

    


#initializing JSON output
result = {"items": [], "variables":{}}


    
    
def main():
    log ("SCRIPT STARTED")
    MY_INPUT = sys.argv[1]
    mySQL = f"SELECT * FROM Mail WHERE Message_ThreadTopic = '{MY_INPUT}' ORDER BY Message_TimeSent ASC"
    handle(mySQL)




def handle(mySQL):
    
    homePath = os.environ['HOME']
    profile = OUTLOOK_DEFAULT_PROFILE

    
    outlookData = homePath + OUTLOOK_DATA_PARENT + profile + OUTLOOK_DATA_FOLDER
    outlookData = "/Users/giovanni/Desktop/Main Profile/Data/"

    MY_DATABASE = f"{outlookData}Outlook.sqlite"
    
    db = sqlite3.connect(MY_DATABASE)
    db.row_factory = sqlite3.Row
    
    rs = db.execute(mySQL).fetchall()
    totCount = len(rs)
    
    myCounter = 0
    attIcon = ''
    for r in rs:
        myCounter += 1
        timeRec = timestampToText (r['Message_TimeSent'])
        if r['Message_HasAttachment'] == 1:
            attIcon = '📎'
        result["items"].append({
            "title": r['Message_NormalizedSubject'],
            
            'subtitle': f"{myCounter}/{totCount:,} {attIcon} {timeRec} From: {r['Message_SenderList']}",
            'valid': True,
            "quicklookurl": '',
            'variables': {
                
            },
                "mods": {

                "control": {
                    "valid": 'true',
                    "subtitle": f"🧵 show entire thread",
                    "arg": r['Message_ThreadTopic']
                },
                "command": {
                    "valid": 'true',
                    "subtitle": f"something with command",
                    "arg": ''
                }},
            "icon": {
                "path": f""
            },
            'arg': r['Message_ThreadTopic']
                }) 

        
    print (json.dumps(result))




            

            
    

if __name__ == '__main__':
    main()
