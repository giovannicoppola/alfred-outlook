#!/usr/bin/env python3
#
# Script to build/update the folder database in outlook
# Created on Saturday, June 3, 2023

from util import log, getFolderData, getAccountData
import json

log ("rebuilding database ⏳...")
getFolderData()
getAccountData()

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



