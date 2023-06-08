#!/usr/bin/env python3
#
# Script to build/update the folder database in outlook
# Created on Saturday, June 3, 2023

from util import log, getFolderData, getAccountData, getContactList, getContactBook
import json
import os

CONTACT_AUTOCOMPLETE = os.getenv('CONTACT_AUTOCOMPLETE')
log ("rebuilding database ⏳...")
getFolderData()
getAccountData()
getContactList()
getContactBook()

log ("done 👍")


result= {"items": [{
        "title": "Done!" ,
        "subtitle": "ready to use outlookSuite now ✅",
        "arg": "",
        "icon": {

                "path": "icons/done.png"
            }
        }]}

print (json.dumps(result))



