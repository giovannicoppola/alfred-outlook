#!/usr/bin/env python3
# encoding: utf-8
# Wednesday, September 1, 2021


from datetime import datetime
import sys
from consts import OUTLOOK_SNOOZER_FILE
import json

with open(OUTLOOK_SNOOZER_FILE, "r") as f:
    myJSON = json.load(f)

if len(sys.argv) >1:
    myInput = str (sys.argv[1])
    
    # Count the number of entries for the user-provided date
    myCount = sum(1 for value in myJSON.values() if value == myInput)

    print (myCount)
	
	
