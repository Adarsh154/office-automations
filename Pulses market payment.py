import urllib.parse
from datetime import datetime, timedelta

import gspread
import pytz
import requests
from oauth2client.service_account import ServiceAccountCredentials

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
TOKEN = "Token_here"
# add credentials to the account
creds = ServiceAccountCredentials.from_json_keyfile_name('storied-fuze-274909-4eea10b19681.json', scope)

# authorize the client sheet
client = gspread.authorize(creds)
# get the instance of the Spreadsheet
sheet = client.open('Market Payment')

# get the first sheet of the Spreadsheet
sheet_instance = sheet.get_worksheet(0)
data = sheet_instance.get_all_values()[1:]


# Function to parse date string to datetime object
def parse_date(date_str):
    naive_date = datetime.strptime(date_str, "%d/%m/%Y")
    # Make it timezone-aware
    aware_date = indian_timezone.localize(naive_date)
    return aware_date


# Get current date
indian_timezone = pytz.timezone('Asia/Kolkata')
current_date = datetime.now(indian_timezone)

# Get the start date of the current week (Monday)
start_of_week = current_date - timedelta(days=current_date.weekday())

# Get the end date of the current week (Sunday)
end_of_week = start_of_week + timedelta(days=6)

# Filter payments due in the current week
due_payments = []
for payment in data:
    due_date = parse_date(payment[5])
    if start_of_week.date() <= due_date.date() <= end_of_week.date() and payment[-2].lower().strip() not in (
            "yes", "y"):
        due_payments.append(payment)

party_wise_payments = {}
total_amount = 0
for payment in due_payments:
    party = payment[1]
    amount = payment[4]
    if party not in party_wise_payments:
        party_wise_payments[party] = []
    party_wise_payments[party].append(payment)
    total_amount += int(amount)

# Print payments party-wise and total amount
text = "Payments party-wise:\n"
for party, payments in party_wise_payments.items():
    text += f"Party: {party}\n"
    party_total = sum(int(payment[4]) for payment in payments)
    for payment in payments:
        readable_list = [payment[0]]
        readable_list.extend(payment[2:5])
        text += ', '.join(readable_list)
        text += "\n"
    text += f"Total amount for {party}: {party_total}\n\n"
text += f"Overall total amount: {total_amount}"
url = "https://api.telegram.org/bot{}/sendMessage?chat_id=chatidhere&text={}".format(
    TOKEN, urllib.parse.quote_plus(text))
requests.get(url)
