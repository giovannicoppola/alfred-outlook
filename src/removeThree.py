#!/usr/bin/env python3

import sys
myString =sys.argv[1]



myString = myString.strip()
if myString.endswith("%0A)"):
    myString = myString[:-4]+')'

print (myString)

