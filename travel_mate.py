import logging
import os
import imaplib
import email
from email.header import decode_header
import re
import pandas as pd
from datetime import datetime
import json

from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font
from collections import defaultdict

# Config Paths (make directories if they don't exist, setup logging)
project_dir = '/Users/prakashkopi/Documents/Personal Code Projects/travel_mate'
log_dir = os.path.join(project_dir, 'logs')
excel_dir= os.path.join(project_dir, 'excel')
os.makedirs(log_dir, exist_ok=True)
os.makedirs(excel_dir, exist_ok=True)

LOG_PATH = os.path.join(log_dir, 'lastRun.log')
UID_FILE_PATH = os.path.join(project_dir, 'uids.txt')
CONFIG_FILE_PATH = os.path.join(project_dir, 'config.json')

logging.basicConfig(filename=LOG_PATH, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Function that writes records to an Excel file
# Each year will have its own file, and each month will have its own sheet within that
def write_to_excel(records):
    if not records:
        logging.info("No records to write to Excel.")
        return

    df = pd.DataFrame(records)

    # Group records by year
    grouped_by_year = defaultdict(list)
    for _, row in df.iterrows():
        year = row["Date"].year
        grouped_by_year[year].append(row)

    for year, year_rows in grouped_by_year.items():
        year_str = str(year)
        filename = f"AA_Travel_Log_{year_str}.xlsx"
        filepath = os.path.join(excel_dir, filename)

        # Load or create the workbook
        if os.path.exists(filepath):
            wb = load_workbook(filepath)
            logging.info("Loading existing file: %s", filepath)
        else:
            wb = Workbook()
            wb.remove(wb.active)  # Remove the default sheet
            logging.info("Creating new file: %s", filepath)

        # Group records by month within the year
        grouped_by_month = defaultdict(list)
        for row in year_rows:
            month_name = row["Date"].strftime("%B %Y")
            grouped_by_month[month_name].append(row)

        for month_name, rows in grouped_by_month.items():
            if month_name not in wb.sheetnames:
                ws = wb.create_sheet(title=month_name)
                ws.append(["Traveler", "Flight No", "From", "To", "Date", "Total Cost (USD)"])
                # Make the header bold
                for cell in ws[1]:
                    cell.font = Font(bold=True)
            else:
                ws = wb[month_name]

            # Append new rows only (don't recreate the sheet)
            for row in rows:
                ws.append([
                    row["Traveler"],
                    row["Flight No"],
                    row["From"],
                    row["To"],
                    row["Date"].strftime("%Y-%m-%d"),
                    row["Total Cost (USD)"]
                ])

            # Recalculate monthly totals per traveler (clear old totals first)
            # Clear existing monthly totals (columns H and beyond)
            for row in ws.iter_rows(min_row=1, min_col=8):
                for cell in row:
                    cell.value = None

            traveler_totals = defaultdict(float)
            for row in ws.iter_rows(min_row=2, max_col=6, values_only=True):
                traveler, _, _, _, _, cost = row
                if isinstance(cost, (int, float)):
                    traveler_totals[traveler] += cost

            start_col = 8
            ws.cell(row=1, column=start_col, value="Traveler").font = Font(bold=True)
            ws.cell(row=1, column=start_col + 1, value="Monthly Total (USD)").font = Font(bold=True)

            for i, (traveler, total) in enumerate(traveler_totals.items(), start=2):
                ws.cell(row=i, column=start_col, value=traveler)
                ws.cell(row=i, column=start_col + 1, value=round(total, 2))

        # Update yearly totals sheet
        if "Yearly Total" not in wb.sheetnames:
            yt = wb.create_sheet(title="Yearly Total")
            yt.append(["Traveler", "Total Cost (USD)"])
            for cell in yt[1]:
                cell.font = Font(bold=True)
        else:
            yt = wb["Yearly Total"]

        yearly_totals = defaultdict(float)
        for sheet_name in wb.sheetnames:
            if sheet_name == "Yearly Total":
                continue
            sheet = wb[sheet_name]
            for row in sheet.iter_rows(min_row=2, max_col=6, values_only=True):
                traveler, _, _, _, _, cost = row
                if isinstance(cost, (int, float)):
                    yearly_totals[traveler] += cost

        # Clear existing yearly totals
        for row in yt.iter_rows(min_row=2):
            for cell in row:
                cell.value = None

        for i, (traveler, total) in enumerate(sorted(yearly_totals.items()), start=2):
            yt.cell(row=i, column=1, value=traveler)
            yt.cell(row=i, column=2, value=round(total, 2))

        # Save workbook
        wb.save(filepath)
        logging.info("Excel file updated: %s", filepath)


# Function that extracts info using regex and returns the actual values
def parseEmail(body):

    # Extract info using regex
    traveler = re.search(r'TRAVELER:\s+(.*)', body)
    flight_no = re.search(r'FLIGHT NO:\s+(\d+)', body)
    dep_station = re.search(r'FROM:\s+(\w+)', body)
    arv_station = re.search(r'TO:\s+(\w+)', body)
    date_flown = re.search(r'DATE FLOWN:\s+([\d/]+)', body)
    total_cost = re.search(r'TOTAL CHARGES BILLED:\s+USD\s+([\d.]+)', body)

    # Extract the actual values from regex matches
    traveler_name = traveler.group(1).strip() if traveler else "Not found"
    flight_number = flight_no.group(1) if flight_no else "Not found"
    departure = dep_station.group(1) if dep_station else "Not found"
    arrival = arv_station.group(1) if arv_station else "Not found"
    flight_date = date_flown.group(1) if date_flown else "Not found"
    total_amount = total_cost.group(1) if total_cost else "Not found"

    # Return the values
    return traveler_name, flight_number, departure, arrival, flight_date, total_amount

# Function to read config file
def read_config():
    with open(CONFIG_FILE_PATH) as f:
        config= json.load(f)
    return config

CONFIG = read_config() # Use throughout the code 

logging.info("--------------- START OF RUN ---------------")

# Load processed UIDs
processed_uids = set()
if os.path.exists(UID_FILE_PATH):
    with open(UID_FILE_PATH, 'r') as f:
        processed_uids= set(f.read().splitlines())
logging.info("Loaded %d processed UIDs from file", len(processed_uids))

# Connect to email server 
mail = imaplib.IMAP4_SSL(CONFIG["Read Email"]["imap_server"])
mail.login(CONFIG["Read Email"]["email_address"], CONFIG["Read Email"]["password"])
mail.select("inbox")

logging.info("Connected successfully to the email server")


# Search for emails from Travel Planner (AA)
status, messages = mail.search(None, '(FROM "NO-REPLY-Travel-Planner_PRD@aa.com")')
email_ids = messages[0].split()
logging.info("Total emails found: %d", len(email_ids))

# Limit how many emails to process (can be changed in config.json)
emailCount= CONFIG["Read Email"]["count"]
logging.info("Email count to process: %d", emailCount)
records = []

skipCount= 0
for eid in email_ids[-emailCount:]:  # Only check the last 10 emails for efficiency
    # Skip already processed emails
    uid = eid.decode()  # Decode bytes to string
    if uid in processed_uids:
        skipCount += 1 # keep track of skipped emails
        continue

    res, msg = mail.fetch(eid, "(RFC822)")
    for response in msg:
        if isinstance(response, tuple):
            msg = email.message_from_bytes(response[1])
            body = ""

            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body += part.get_payload(decode=True).decode()
            else:
                body = msg.get_payload(decode=True).decode()


            # Extract info using regex and return the actual values
            traveler_name, flight_number, departure, arrival, flight_date, total_amount= parseEmail(body)

            # If the email has all the necessary info, append it to records
            if all([traveler_name, flight_number, departure, arrival, flight_date, total_amount]):
                try:
                    record = {
                        "Traveler": traveler_name,
                        "Flight No": flight_number,
                        "From": departure,
                        "To": arrival,
                        "Date": datetime.strptime(flight_date, "%m/%d/%Y"),
                        "Total Cost (USD)": float(total_amount)
                    }
                    records.append(record)
                    logging.info("Parsed and added record: %s", record)
                    processed_uids.add(uid)
                except Exception as e:
                    logging.error("Error processing email UID %s: %s", uid, str(e))

# Save to excel file
write_to_excel(records)

# Send invoice to each traveler 



# Save updated UID list
with open(UID_FILE_PATH, 'w') as f:
    for uid in processed_uids:
        f.write(f"{uid}\n")

logging.info("Total records skipped because they were already processed: %d", skipCount)
logging.info("Total new records added: %d", len(records))

logging.info("--------------- END OF EXECUTION ---------------")
print("--------------- END OF EXECUTION ---------------")