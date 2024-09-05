import yaml
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
import re
import os

with open("settings_path.txt") as f:
    settings_path = f.read()

print("Settings path:",settings_path)

if os.path.exists(settings_path):
    # Load settings from YAML file
    with open(settings_path) as stream:
        try:
            formatted_params = yaml.safe_load(stream)
        except yaml.YAMLError as exc:
            print(exc)
else:
    print(f"{settings_path} not found")

print(formatted_params)

# Function to authenticate Google Sheets client
def authenticate_client():
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)
    return client

# Function to authenticate and return a specific worksheet
def authenticate_sheet(sheet_url, worksheet_name):
    client = authenticate_client()
    sheet = client.open_by_url(sheet_url).worksheet(worksheet_name)
    return sheet

# Function to filter the sheet based on specific criteria
def get_filtered_sheet(sheet_url, worksheet_name, target_subsystem):
    
    def find_and_check_date(string):
        date_pattern = r'\b(\d{1,2}/\d{1,2}/\d{4})\b'
        match = re.search(date_pattern, string)
        if match:
            return True
        return False

    def validate_notification_field(string):
        return (string[:len('notified')].lower() == 'notified') and find_and_check_date(string)

    def validate_subsystem_field(string):
        return string.lower().strip() == target_subsystem.lower().strip()

    sheet = authenticate_sheet(sheet_url, worksheet_name)
    data = sheet.get_all_values()
    df = pd.DataFrame(data[1:], columns=data[0])

    preference_columns = [formatted_params["columns"]["preference1"], formatted_params["columns"]["preference2"]]
    if target_subsystem not in [None, ""]:
        filtered_df = df[df[preference_columns[0]].apply(validate_subsystem_field) | 
                         df[preference_columns[1]].apply(validate_subsystem_field)]    

    if "Notified_" + target_subsystem in filtered_df.keys():
        filtered_df = filtered_df[filtered_df["Notified_" + target_subsystem].apply(validate_notification_field)]

    # Extract 'id', 'Interview Date', and 'Interview Time' from the 'Notified_' column
    notified_column = "Notified_" + target_subsystem
    filtered_df['id'] = filtered_df.index + 2  # Add 2 to match the row number in the sheet (considering header)
    filtered_df['Interview Date'] = filtered_df[notified_column].apply(lambda x: re.search(r'\b(\d{1,2}/\d{1,2}/\d{4})\b', x).group() if re.search(r'\b(\d{1,2}/\d{1,2}/\d{4})\b', x) else '')
    filtered_df['Interview Time'] = filtered_df[notified_column].apply(lambda x: re.search(r'(\d{1,2}:\d{2})', x).group() if re.search(r'(\d{1,2}:\d{2})', x) else '')

    # Convert date and time to datetime for sorting
    try:
        filtered_df['Interview Datetime'] = pd.to_datetime(filtered_df['Interview Date'] + ' ' + filtered_df['Interview Time'], format='%d/%m/%Y %H:%M', errors='coerce')
    except Exception as e:
        print("Error in datetime conversion:", e)
        print(filtered_df[['Interview Date', 'Interview Time']])  # Debugging info
    
    # Check if 'Interview Datetime' was created
    if 'Interview Datetime' not in filtered_df.columns:
        raise KeyError("Interview Datetime column was not created successfully.")

    # Ensure 'id' is the first column in the DataFrame
    columns = ['id'] + [col for col in filtered_df.columns if col != 'id']
    filtered_df = filtered_df[columns]

    return filtered_df

# Function to append filtered data to the sheet after sorting and checking for duplicates
def append_df_to_sheet(df, sheet_url, worksheet_name):
    client = authenticate_client()
    sheet = client.open_by_url(sheet_url)
    
    try:
        worksheet = sheet.worksheet(worksheet_name)
        print(f"Worksheet '{worksheet_name}' found.")
    except gspread.exceptions.WorksheetNotFound:
        # If the worksheet doesn't exist, create it
        num_columns = len(df.columns)
        worksheet = sheet.add_worksheet(title=worksheet_name, rows="1000", cols=num_columns)
        headers = df.columns.tolist()  # Headers with 'id' as the first column
        worksheet.update('A1', [headers])
        print(f"Worksheet '{worksheet_name}' created with {num_columns} columns.")

    # Ensure sorting by Interview Datetime before appending
    if 'Interview Datetime' in df.columns:
        df = df.sort_values(by='Interview Datetime')
        # Drop the 'Interview Datetime' column before appending
        df = df.drop(columns=['Interview Datetime'])
    else:
        print("'Interview Datetime' column not found for sorting.")

    # Get existing data from the worksheet
    existing_data = worksheet.get_all_values()
    if len(existing_data) > 1:  # More than just headers
        existing_df = pd.DataFrame(existing_data[1:], columns=existing_data[0])
        
        # Convert 'id' column to numeric for comparison
        existing_df['id'] = pd.to_numeric(existing_df['id'], errors='coerce')
        df['id'] = pd.to_numeric(df['id'], errors='coerce')

        # Identify new rows by checking 'id' column
        new_rows = df[~df['id'].isin(existing_df['id'])]

        if not new_rows.empty:
            # Append new rows to the worksheet
            for _, row in new_rows.iterrows():
                worksheet.append_row(row.tolist(), value_input_option='RAW')  # Ensure 'id' is the first column
            print(f"Appended {len(new_rows)} new rows of data.")
        else:
            print("No new rows to add.")
    else:
        # If the sheet is empty or has only headers, append all data
        values = df.values.tolist()
        for row in values:
            worksheet.append_row(row, value_input_option='RAW')
        print(f"Added {len(df)} rows of data.")

    return sheet_url

# Define the columns to be used
columns = formatted_params['columns']
column_aliases = formatted_params['columns'].keys()
columns = [formatted_params['columns'][alias] for alias in column_aliases if alias not in ['first_year', 'notifier']]
tempcolumns = ['id', 'Interview Date', 'Interview Time', 'Interview Datetime']
for column in columns:
    tempcolumns.append(column)
columns = tempcolumns

technical_columns = ["Interviewer 1","Interviewer 2","Technical Score 1", "Technical Score 2", "Ethu Score 1", "Enthu Score 2", 
                "Technical average", "Verdict 1", "Verdict 2", "Additional Comments", "Cock Score"]

sna_columns = ["Interviewer", "Technical Score", "Enthu Score", "Technical Average", "Additional Comments", "Cock Score"]

management_columns = ["Communication (3)", "Aptitude(2)", "Creativity(3)" ,"Graphics(2)", "Videography/Webdev(3)", "Admin", "Finance/PR(2)","Cock Score","TOTAL(10)","Additional Comments"]

subsystems = {
    "Artificial Intelligence": technical_columns,
    "Sensing And Automation": sna_columns,
    "Management": management_columns,
    "Mechanical": technical_columns
}

for subsystem in subsystems.keys():
    # Get filtered data from the target sheet and worksheet
    df = get_filtered_sheet(formatted_params['sheet_settings']['target_sheet_url'], formatted_params['sheet_settings']['target_worksheet_name'], subsystem)[columns]

    # Add additional columns with empty values
    new_columns = subsystems[subsystem]
    for column in new_columns:
        df[column] = len(df) * ['']

    # Append the DataFrame to the target worksheet
    append_df_to_sheet(df, formatted_params['sheet_settings']['interview_score_url'], subsystem)
