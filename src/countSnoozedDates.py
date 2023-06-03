#!/usr/bin/env python
# encoding: utf-8
# Wednesday, May 5, 2021
# added number of days on Thursday, February 17, 2022, 2:35 PM
# updated and refreshed on Friday, June 2, 2023

from consts import OUTLOOK_SNOOZER_FILE
from datetime import date, datetime, timedelta
import json

with open(OUTLOOK_SNOOZER_FILE, "r") as f:
    myJSON = json.load(f)

# Get the current date
today = date.today()

# Create an empty report string
report = ""

# Sort the dictionary based on the date values
sorted_dict = sorted(myJSON.items(), key=lambda x: x[1])

# Create a dictionary to store the count for each unique date
date_counts = {}

# Iterate over the sorted key-value pairs
for key, value in sorted_dict:
    # Convert the date string to a datetime object
    date_obj = datetime.strptime(value, "%Y-%m-%d").date()

    # Calculate the number of days from today
    days_difference = (date_obj - today).days

    # Format the date as "Weekday, Month Day"
    formatted_date = date_obj.strftime("%A, %B %d")

    # Add or increment the count for the date in the dictionary
    date_counts.setdefault(formatted_date, 0)
    date_counts[formatted_date] += 1

# Construct the report string with one line per date
for formatted_date, count in date_counts.items():
    days_difference = (datetime.strptime(formatted_date + f", {today.year}", "%A, %B %d, %Y").date() - today).days
    report += f"{formatted_date} ({days_difference}): {count}\n"

# Print the report string
print(report)
	
	
