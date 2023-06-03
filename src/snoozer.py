"""
A new version of the OL SNOOZER script, to be included in the alfred-OutlookSuite
Sunny ☀️   🌡️+70°F (feels +70°F, 59%) 🌬️↓2mph 🌕&m Fri Jun  2 08:42:48 2023
W22Q2 – 153 ➡️ 211 – 22 ❇️ 343
"""

"""
# Script updated on Mon Apr 11 10:17:50 2022  – Sunny ☀️   🌡️+48°F (feels +45°F, 39%) 🌬️↙9mph 🌔 
# to use exchange id instead of database ID, so that it will be possible to use it across computers. 
# Partly cloudy ⛅️  🌡️+69°F (feels +69°F, 41%) 🌬️↘13mph 🌔 Tue Apr 12 17:03:32 2022
# tried to fetch the message ID from the sqlite database. To my knowledge is the only way to obtain a message ID which is shared across computers. The unique id I have been using is not. 
# it seems there is no way to get this via applescript

# also, removed the categories (labels) part which I don't use for now. These are in v01 of this script, in the same folder. 
"""
from subprocess import Popen, PIPE, check_output
import sys
from util import log, fetchRecordID
import json
from consts import *


def main():

    MY_SNOOZE_DATE = sys.argv[1]
    

    scpt = '''
        tell application "Microsoft Outlook"
            
            activate
            
            if view of the first main window is not equal to "mail view" then
                set view of the main window 1 to mail view
                
            end if
            
            set msgSet to selection
            
            set myIDs to {}
            
            repeat with aMessage in msgSet
                
                set msgID to id of aMessage # getting regular ID (which is different across computers)		
                set end of myIDs to msgID
                move aMessage to folder "Snoozed"
                
                
                
            end repeat
            
        end tell

        set AppleScript's text item delimiters to ":::"
        set myIDs to myIDs as text
        set AppleScript's text item delimiters to ""
        return myIDs
        
    end run '''

    command = ['osascript', '-e', scpt]
    myOutput = check_output(command).decode('utf-8').strip()


    
    myIDs = myOutput.split(":::")
    
    if os.path.exists(OUTLOOK_SNOOZER_FILE):
        with open(OUTLOOK_SNOOZER_FILE, "r") as f:
            mySnoozes = json.load(f)
    else:
        mySnoozes = {}

    for myID in myIDs:
        myRecordID = fetchRecordID (myID)
        
        mySnoozes[myRecordID] = MY_SNOOZE_DATE
    
    with open(OUTLOOK_SNOOZER_FILE, "w") as f:
        json.dump(mySnoozes, f, indent=4)

    log ("Snoozer file updated")         
    



if __name__ == '__main__':
    main()


