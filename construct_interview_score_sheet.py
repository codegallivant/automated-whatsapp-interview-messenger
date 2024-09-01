import yaml
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import datetime
import gspread
from google.auth.exceptions import RefreshError
import re


with open("settings.yaml") as stream:
    try:
        PARAMS = yaml.safe_load(stream)
        print()
    except yaml.YAMLError as exc:
        print(exc)


def authenticate_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    return client


def authenticate_sheet(sheet_url, worksheet_name):
    client = authenticate_client()
    sheet = client.open_by_url(sheet_url).worksheet(worksheet_name)
    return sheet


def get_filtered_sheet(sheet_url, worksheet_name):
    
    def find_and_check_date(string):
        # Regular expression to match dates in DD/MM/YYYY format
        date_pattern = r'\b(\d{2}/\d{2}/\d{4})\b'
        
        # Search for the date in the string
        match = re.search(date_pattern, string)
        
        if match:
            date_str = match.group(1)
            date_obj = datetime.datetime.strptime(date_str, '%d/%m/%Y')
            today = datetime.datetime.today()
            if date_obj.date() == today.date():
                return True
        return False

    def validate_notification_field(string):
        return (string[:len('notified')].lower() == 'notified') and find_and_check_date(string)

    def validate_subsystem_field(string):
        return string.lower().strip() == PARAMS["subsystem_settings"]["target_subsystem"].lower().strip()

    sheet = authenticate_sheet(sheet_url, worksheet_name)

    data = sheet.get_all_values()

    df = pd.DataFrame(data[1:], columns=data[0])
    
    preference_columns = [PARAMS["columns"]["preference1"], PARAMS["columns"]["preference2"]]
    if PARAMS["subsystem_settings"]["target_subsystem"] not in [None, ""]:
        filtered_df = df[df[preference_columns[0]].apply(validate_subsystem_field) | 
                        df[preference_columns[1]].apply(validate_subsystem_field)]    
    filtered_df = filtered_df[filtered_df[PARAMS["columns"]["first_year"]]=="Yes"]

    # print(filtered_df["Notified_"+PARAMS["subsystem_settings"]["target_subsystem"]].apply(validate_notification_field))
    if "Notified_"+PARAMS["subsystem_settings"]["target_subsystem"] in filtered_df.keys():
        filtered_df = filtered_df[filtered_df["Notified_"+PARAMS["subsystem_settings"]["target_subsystem"]].apply(validate_notification_field)]

    filtered_df['id'] = filtered_df.index + 2

    return filtered_df


def append_df_to_sheet(df, sheet_url, worksheet_name):
    client = authenticate_client()

    try:
        # Try to open the existing sheet
        sheet = client.open_by_url(sheet_url)
        print(f"Sheet found: {sheet.title}")
    except (gspread.exceptions.SpreadsheetNotFound, RefreshError):
        # If the sheet doesn't exist, create it
        sheet = client.create(f"New Sheet for {worksheet_name}")
        print(f"New sheet created: {sheet.title}")
        
        # Share the sheet with the user's email (replace with actual email)
        sheet.share('user@example.com', perm_type='user', role='writer')
        
        # Update the sheet_url with the new sheet's URL
        sheet_url = f"https://docs.google.com/spreadsheets/d/{sheet.id}"
        print(f"New sheet URL: {sheet_url}")

    try:
        # Try to open the worksheet
        worksheet = sheet.worksheet(worksheet_name)
        print(f"Worksheet '{worksheet_name}' found.")
    except gspread.exceptions.WorksheetNotFound:
        # If the worksheet doesn't exist, create it with the DataFrame columns
        num_columns = len(df.columns)
        worksheet = sheet.add_worksheet(title=worksheet_name, rows="1000", cols=num_columns)
        
        # Add column headers
        headers = df.columns.tolist()
        worksheet.update('A1', [headers])
        
        print(f"Worksheet '{worksheet_name}' created with {num_columns} columns.")

    existing_data = worksheet.get_all_values()

    if len(existing_data) == 1:  # Only headers exist
        values = df.values.tolist()
        worksheet.append_rows(values)
        print(f"Added {len(df)} rows of data.")
    else:
        existing_header = existing_data[0]
        df_header = df.columns.tolist()

        if existing_header != df_header:
            print("Warning: Existing header does not match DataFrame columns.")
            print(f"Existing header: {existing_header}")
            print(f"DataFrame columns: {df_header}")
            return

        values = df.values.tolist()
        worksheet.append_rows(values)
        print(f"Appended {len(df)} rows of data.")

    return sheet_url


columns = PARAMS['columns']
column_aliases = PARAMS['columns'].keys()
columns = [PARAMS['columns'][alias] for alias in column_aliases if alias not in ['first_year','notifier']]

df = get_filtered_sheet(PARAMS['sheet_settings']['target_sheet_url'],PARAMS['sheet_settings']['target_worksheet_name'])

df = df[columns]
print(df)

new_columns = ["Technical Score 1", "Technical Score 2", "Ethu Score 1", "Enthu Score 2", "Technical average", "Verdict 1", "Verdict 2", "Additional Comments"]
for column in new_columns:
    df[column] = len(df)*['']

append_df_to_sheet(df, PARAMS['sheet_settings']['interview_score_url'],PARAMS['subsystem_settings']['target_subsystem'])


