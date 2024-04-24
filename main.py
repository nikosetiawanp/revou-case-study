# Prompt
import calendar
import questionary

# env
import os
from dotenv import load_dotenv

# Data Fetching
import sys
import gspread
import requests
import json
from gspread_formatting import format_cell_ranges
from google.oauth2.service_account import Credentials
from datetime import datetime, timedelta

import pytz
from collections import defaultdict


load_dotenv()

day_names = ["Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu", "Minggu"]
month_names = [
    "Januari",
    "Februari",
    "Maret",
    "April",
    "Mei",
    "Juni",
    "Juli",
    "Agustus",
    "September",
    "Oktober",
    "November",
    "Desember",
]
month_map = {name: index + 1 for index, name in enumerate(month_names)}


def unix_timestamp_to_date(unix_timestamp):
    # Convert Unix timestamp to datetime object
    dt = datetime.fromtimestamp(unix_timestamp)

    # Format datetime object
    formatted_date = dt.strftime("%d %B %Y")

    return formatted_date


def get_weekly_date_range_unix_timestamp(year, week_number):
    # Calculate the start date of the week (Monday)
    start_date = datetime(year, 1, 1) + timedelta(
        weeks=week_number - 1, days=-datetime(year, 1, 1).weekday()
    )
    start_date = start_date.replace(hour=0, minute=0, second=0, microsecond=0)

    # Calculate the end date of the week (Sunday)
    end_date = start_date + timedelta(days=6, hours=23, minutes=59, seconds=59)
    end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=0)

    # Convert dates to Unix timestamps
    start_timestamp = int(start_date.timestamp())
    end_timestamp = int(end_date.timestamp())

    return start_timestamp, end_timestamp


def get_current_week_number():
    # Get the current date
    today = datetime.now()

    # Get the ISO week number (1-53) and the weekday (1-7, where Monday is 1 and Sunday is 7)
    week_number = today.strftime("%V")

    return int(week_number)

def unix_timestamp_to_custom_format(unix_timestamp):
    # Convert Unix timestamp to datetime object
    dt = datetime.fromtimestamp(unix_timestamp)

    # Get the day of the week (0 for Monday, 6 for Sunday)
    day_of_week = dt.weekday()

    # Get the day, month, and year
    day = dt.day
    month = month_names[dt.month - 1]
    year = dt.year

    # Construct the formatted string
    formatted_date = f"{day_names[day_of_week]}, {day} {month} {year}"

    return formatted_date

def get_weeks_in_month(year, month):
    num_days_in_month = calendar.monthrange(year, month)[1]
    num_weeks = (num_days_in_month + calendar.weekday(year, month, 1) + 6) // 7

    return num_weeks

def get_week_number(year, month, week):
    first_day_of_month = datetime(year, month, 1)
    days_to_add = (int(week) - 1) * 7 
    target_date = first_day_of_month + timedelta(days=days_to_add)
    week_number = target_date.isocalendar()[1]

    return week_number

now = datetime.now()
current_week_number = get_current_week_number()
current_year_number = now.year

# Prompts

# Select year
selected_year_number = int(
    questionary.select(
        "Select year:",
        choices=["2024", "2023", "2022", "2021", "2020", "2019"],
    ).ask()
)

# Input Week
selected_week_number = int(
    input(f"Please enter the week number (Today: 2024 Week {current_week_number}): ")
)

created_after = get_weekly_date_range_unix_timestamp(selected_year_number, selected_week_number)[
    0
]
created_before = get_weekly_date_range_unix_timestamp(
    selected_year_number, selected_week_number
)[1]

# Huntr
url = os.getenv("API_URL")
token = os.getenv("ACCESS_TOKEN")
headers = {"Authorization": f"Bearer {token}"}

# Fetch data from huntr
def get_activities_list():
    print(f"Fetching activities from {unix_timestamp_to_custom_format(created_after)} to {unix_timestamp_to_custom_format(created_before)}")
    activities = []
    next_token_list = []

    def fetch():
        response = requests.get(
            url + "activities",
            headers=headers,
            params={
                "limit": 500,
                "created_before": created_before,
                "created_after": created_after,
            },
        )
        if response.status_code == 200:
            response_data = json.loads(response.text)
            activities.extend(response_data["data"])
            print(f"Found {len(activities)} total activities")
            next_token_list.append(response_data["next"])
        else:
            print(f"Error: {response.status_code} - {response.reason}")

    def fetch_next():
        response = requests.get(
            url + "activities",
            headers=headers,
            params={
                "limit": 500,
                "created_before": created_before,
                "created_after": created_after,
                "next": next_token_list[len(next_token_list) - 1],
            },
        )
        if response.status_code == 200:
            response_data = json.loads(response.text)
            activities.extend(response_data["data"])
            print(f"Found {len(activities)} total activities")
            next_token_list.append(response_data["next"])
        else:
            print(f"Error: {response.status_code} - {response.reason}")

    # Continuously fetch data while there's next page
    fetch()
    while next_token_list[len(next_token_list) - 1] is not None:
        fetch_next()

    # Simplifying data
    data = [
        {
            "email": activity["ownerMember"]["email"],
            "createdAt": activity["createdAt"],
            "activity": activity["activityCategory"]["name"],
        }
        for activity in activities
    ]
    sorted_data = sorted(data, key=lambda x: (x["email"], x["activity"]))
    return sorted_data


data = get_activities_list()  # data from huntr
activity_counts = defaultdict(lambda: defaultdict(int))  # store count of each activity

if not data:
    print("No activity found, aborting process")
    sys.exit()  # Abort the script

# Count activities of each email
for entry in data:
    email = entry["email"]
    activity = entry["activity"]
    activity_counts[email][activity] += 1

# Grouping each activity of each email
grouped_data = [
    {"email": email, **activities} for email, activities in activity_counts.items()
]

# Types of activities
activity_types = [
    "Apply",
    "Create Cover Letter",
    "Research Company / Job Requirement",
    "Priority Job",
    "Upload your CV in the document sections",
    "Upload your cover letter in the document sections",
    "Received User Invitation",
    "Rejected",
    "Networking Event",
    "Accept Offer",
]

# Preparing grouped data to push to gspread
data_to_push = []
for entry in grouped_data:
    mapped_entry = [entry["email"]]
    for activity_type in activity_types:
        if activity_type in entry:
            mapped_entry.append(entry[activity_type])
        else:
            mapped_entry.append(0)
    data_to_push.append(mapped_entry)

# Google sheets API
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# Gspread
sheet_id = os.getenv("SHEET_ID")
workbook = client.open_by_key(sheet_id)
worksheet_list = map(lambda x: x.title, workbook.worksheets())  # Get all sheet title
new_worksheet_name = f"{selected_year_number} W{selected_week_number}"

if len(data_to_push) == 0:
    print("No data to push, aborting process")
    
# Create worksheet if doesn't exist, otherwise write into that worksheet
if new_worksheet_name in worksheet_list:
    worksheet = workbook.worksheet(new_worksheet_name)
else:
    print(f"Creating {selected_year_number} W{selected_week_number} sheet")
    worksheet = workbook.add_worksheet(new_worksheet_name, rows=10, cols=10)

# Pushing to gspread
table_headers = [
    "Email",
    "Apply",
    "Create Cover Letter",
    "Research Company",
    "Priority Job",
    "Upload CV",
    "Upload cover letter",
    "Received User Invitation",
    "Rejected",
    "Networking Event",
    "Accept Offer",
]

try:
    print(f"Updating {selected_year_number} W{selected_week_number} sheet")
    worksheet.clear()  # Clear existing data
    worksheet.update([[f"{selected_year_number} W{selected_week_number}"]], "A1")
    worksheet.update(
        [
            [
                f"{unix_timestamp_to_custom_format(created_after)} - {unix_timestamp_to_custom_format(created_before)}"
            ]
        ],
        "C1",
    )
    worksheet.update([table_headers], "A2")  # Push headers to gspread
    worksheet.update(data_to_push, "A3")  # Push data to gspread
    worksheet.format(
        "A1:K1",
        {
            "textFormat": {"bold": True},
        },
    )

    worksheet.format("A2:K2", {"textFormat": {"bold": True}})
    worksheet.format("B2:K2", {"horizontalAlignment": "RIGHT"})

    print(f"Successfully updated '{selected_year_number} W{selected_week_number}' sheet")
except:
    print(
        f"Failed updating '{selected_year_number} W{selected_week_number}' sheet, please try again"
    )
