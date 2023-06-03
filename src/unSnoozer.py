"""
A new version of the OL SNOOZER script, to be included in the alfred-OutlookSuite
Sunny ☀️   🌡️+70°F (feels +70°F, 59%) 🌬️↓2mph 🌕&m Fri Jun  2 08:42:48 2023
W22Q2 – 153 ➡️ 211 – 22 ❇️ 343
"""

# Snooze by date applescript, Monday, May 25, 2020, 1:17 PM
# modified on Saturday, May 30, 2020, 1:25 PM to overwrite the ID and date files and avoid duplications
## Thursday, November 19 2020, 9:45AM added count of unsnoozing tomorrow, which is helpful to plan next day
## Friday, Decemeber 11 2020, added multiple days

# on Mon Apr 11 10:17:50 2022  Sunny ☀️   🌡️+48°F (feels +45°F, 39%) 🌬️↙9mph 🌔 
# tried to use exchange id instead of database ID, so that it will be possible to use it across computers, however the exchange ID changes if the message changes folder. 

# Partly cloudy ⛅️  🌡️+69°F (feels +69°F, 41%) 🌬️↘13mph 🌔 Tue Apr 12 17:03:32 2022
# tried to fetch the message ID from the sqlite database. To my knowledge is the only way to obtain a message ID which is shared across computers. The unique id I have been using is not. 
#### I copied the applescript wihtin alfred as opposed to run it from the bash because I don't have sudo permissions to make the applescript executable. 

# Sunny ☀️   🌡️+87°F (feels +86°F, 35%) 🌬️←7mph 🌕&m Fri Jun  2 13:04:58 2023
#W22Q2 – 153 ➡️ 211 – 22 ❇️ 343
# streamlined version to be included in the alfred-outlookSuite workflow


from subprocess import Popen, PIPE, check_output
from util import log, fetchEmailID
from consts import *
import json
from datetime import date


def main ():
    # Get the current date
    today = date.today()

    # Convert today's date to the same format as the dictionary values
    search_date = today.strftime("%Y-%m-%d")
    
    #reading in the JSON file with the snooze information
    if os.path.exists(OUTLOOK_SNOOZER_FILE):
        with open(OUTLOOK_SNOOZER_FILE, "r") as f:
            mySnoozes = json.load(f)
    else:
        mySnoozes = {}
    
    
    # Use list comprehension to filter keys based on the search date or earlier
    myuniIDs = [key for key, value in mySnoozes.items() if value <= search_date]
    #log (myuniIDs) #callling uniIDs those unique ones stored in the JSON file
    
    
    
    # Create a new dictionary with entries not matching the search date or earlier
    new_snooze = {key: value for key, value in mySnoozes.items() if value > search_date}

    with open(OUTLOOK_SNOOZER_FILE, "w") as f:
        json.dump(new_snooze, f, indent=4)

    log ("Snoozer file updated")         
    

    #converting them into record_IDs
    myIDs = []
    for myuniID in myuniIDs:
        #log (myuniID)
        myEmailID = fetchEmailID (myuniID)
        myIDs.append (myEmailID)
    
        
    
    scpt = '''

    on run {myIDs}
        set AppleScript's text item delimiters to ","
        set IDlist to every text item of myIDs
        set AppleScript's text item delimiters to {""} -- Reset delimiters

        tell application "Microsoft Outlook"
            set mailAccount to first exchange account
            set myMailbox to folder "Snoozed" of mailAccount
            set theMsgs to (every message of myMailbox)
            
            repeat with aMsg in theMsgs -- going through all the messages in the Snoozed folder
                set msgID to (id of aMsg) as string
                
                repeat with i in IDlist
                    if msgID is (i as string) then
                        #log "found one"
                        move aMsg to folder "Inbox" of mailAccount
                    end if
                end repeat
            end repeat
        end tell
    end run
            '''

    IDString = ",".join(myIDs) #passing the list of IDs as a string
    command = ['osascript', '-e', scpt, IDString]
    check_output(command).decode('utf-8').strip()


if __name__ == '__main__':
    main()

"""	



	
	
	
	####BEEMINDER section
	#beeminder POST syntax:
	#data = [
	#  ('auth_token', 'YOUR-BEEMINDER-AUTH-TOKEN-GOES-HERE'),
	#  ('value', currentNoteCount),
	#  ('comment', 'from EvernoteScript!'),
	#]
	
	set InboxCount to (count messages in folder "Inbox" of mailAccount)
	set SnoozeCount to (count messages in folder "Snoozed" of mailAccount)
	
	set theURL to "https://www.beeminder.com/api/v1/users/giovanni/goals/OutlookZero/datapoints.json" & " -d " & "auth_token=vpQ9Rm5S24HKysggZkSt -d " & "value=" & InboxCount & " -d " & "comment=from+snoozeScript"
	#	log quoted form of the theURL
	
end tell

# posting to beeminder board
do shell script "curl -X POST " & theURL


#logging 
set LogText to ("ToInbox=" & toInboxCounter & ", Snoozed=" & SnoozeCount & ", TotalInbox=" & InboxCount)

#log LogText

set myLogPath to ("/Users/giovanni.coppola/OneDrive - Regeneron Pharmaceuticals, Inc/MyScripts/SnoozeMail/snoozeLog.txt")
do shell script "echo " & LogText & ", $(date) >> " & quoted form of myLogPath

display dialog ((current date) as string) & "
" & LogText & "

" & "Unsnoozing tomorrow=" & TomorrowCounter & "
" & myDate2_day & ": " & myDate2_c & "
" & myDate3_day & ": " & myDate3_c & "
" & myDate4_day & ": " & myDate4_c


"""
