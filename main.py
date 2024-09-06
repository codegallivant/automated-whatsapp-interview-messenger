import pandas as pd
import yaml
import traceback
import datetime
import time
from alright import WhatsApp
import sys
import multiprocessing
import time
import gspread
import pandas as pd
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import re
import os


def permission_to_continue():
    while True:
        print()
        response = input("Continue execution? (y/n): ")
        print()
        if response in ["y","Y"]:
            return
        elif response in ["n","N"]:
            sys.exit()    


def print_params(params):
    print()
    for param in params.keys():
        if isinstance(params[param], dict):
            print()
            print(param)
            print_params(params[param])
        else:
            print(param, ":", params[param])


with open("settings_path.txt") as f:
    settings_path = f.read()

print("Settings path:", settings_path)

if os.path.exists(settings_path):
    with open(settings_path) as stream:
        try:
            formatted_params = yaml.safe_load(stream)
            print_params(formatted_params)
            print()
        except yaml.YAMLError as exc:
            print(exc)
else:
    print(f"{settings_path} not found")
    sys.exit()

PARAMS = dict()
PARAMS["columns"] = formatted_params["columns"]
PARAMS["testing"] = formatted_params["testing"]
for param in formatted_params.keys():
    if param not in ["columns","testing"]:
        if isinstance(formatted_params[param], dict):
            for key in formatted_params[param].keys():
                PARAMS[key] = formatted_params[param][key]
        else:
            PARAMS[param] = formatted_params[param]

permission_to_continue()

def chrome_options():
    chrome_options = Options()
    if sys.platform == "win32":
        if PARAMS["headless"] == True:
            chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("--profile-directory=Default")
        chrome_options.add_argument("--user-data-dir=C:/Temp/ChromeProfile")
    else:
        if PARAMS["headless"] == True:
            chrome_options.add_argument("--headless=new")
        chrome_options.add_argument("start-maximized")
        chrome_options.add_argument(f"--user-data-dir=./user_data/{PARAMS['notifier']}_User_Data")
    chrome_options.binary_location = PARAMS["path_to_chrome"]
    chrome_options.executable_path= PARAMS["path_to_chromedriver"]
    return chrome_options


def authenticate_sheet(sheet_url, worksheet_name):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    client = gspread.authorize(creds)

    sheet = client.open_by_url(sheet_url).worksheet(worksheet_name)

    return sheet


def get_filtered_sheet():
    
    def validate_notification_field(string):
        return (string[:len('notified')].lower() != 'notified') and (string != 'Message timed out')

    def validate_subsystem_field(string):
        return string.lower().strip() == PARAMS["target_subsystem"].lower().strip()

    sheet_url = PARAMS["target_sheet_url"]
    worksheet_name = PARAMS["target_worksheet_name"]

    sheet = authenticate_sheet(sheet_url, worksheet_name)

    data = sheet.get_all_values()

    filtered_df = pd.DataFrame(data[1:], columns=data[0])

    if PARAMS["check_notifier_column"] == True:
        filtered_df = filtered_df[filtered_df[PARAMS["columns"]["notifier"]] == PARAMS["notifier"]]
    
    filtered_df = filtered_df[filtered_df[PARAMS["columns"]["first_year"]]=="Yes"]

    preference_columns = [PARAMS["columns"]["preference1"], PARAMS["columns"]["preference2"]]
    if PARAMS["target_subsystem"] not in [None, ""]:
        filtered_df = filtered_df[filtered_df[preference_columns[PARAMS["subsystem_preference"]-1]].apply(validate_subsystem_field)]
        # filtered_df = filtered_df[filtered_df[preference_columns[PARAMS["subsystem_preference"]-1]]==PARAMS["target_subsystem"]]
    
    if "Notified_"+PARAMS["target_subsystem"] in filtered_df.keys():
        filtered_df = filtered_df[filtered_df["Notified_"+PARAMS["target_subsystem"]].apply(validate_notification_field)]

    filtered_df['id'] = filtered_df.index + 2

    t1 = datetime.datetime.strptime(PARAMS["start_time"], "%H:%M")
    t2 = datetime.datetime.strptime(PARAMS["end_time"], "%H:%M")
    diff = t2 - t1
    minutes = diff.total_seconds() / 60
    available_time = abs(int(minutes))

    interview_count = (available_time//PARAMS["duration"])*PARAMS["at_once"]
    filtered_df = filtered_df.head(interview_count)

    notified_column = "Notified_" + PARAMS['target_subsystem']

    return filtered_df


details = get_filtered_sheet()
interview_sheet = authenticate_sheet(PARAMS["interview_score_url"], PARAMS["target_subsystem"])
interview_sheet_data = interview_sheet.get_all_values()
interview_sheet_data = pd.DataFrame(interview_sheet_data[1:], columns=interview_sheet_data[0])
interview_sheet_data['id'] = pd.to_numeric(interview_sheet_data['id'], errors='coerce')
details['id'] = pd.to_numeric(details['id'], errors='coerce')
columns = formatted_params['columns']
column_aliases = formatted_params['columns'].keys()
columns = [formatted_params['columns'][alias] for alias in column_aliases if alias not in ['first_year', 'notifier']]
tempcolumns = ['id', 'Interview Date', 'Interview Time']
for column in tempcolumns:
    if column not in details.keys():
        details[column] = ['']*len(details)
for column in columns:
    tempcolumns.append(column)
columns = tempcolumns
details = details[columns]
new_rows = details[~details['id'].isin(interview_sheet_data['id'])]


def append_row_to_sheet(interview_worksheet, source_row_index, interview_time):
    global new_rows
    this_rows = new_rows[new_rows['id']==source_row_index]
    this_rows.loc[:,'Interview Date'] = PARAMS['date']
    this_rows.loc[:,'Interview Time'] = interview_time
    if not this_rows.empty:
        for _, row in this_rows.iterrows():
            interview_worksheet.append_row(row.tolist(), value_input_option='RAW')  # Ensure 'id' is the first column
        print(f"Appended {len(this_rows)} new rows of data.")
    else:
        print("No new rows to add.")

def update_sheet_values(sheet_url, worksheet_name, update_column, row_indexes, update_values):
    sheet = authenticate_sheet(sheet_url, worksheet_name)

    headers = sheet.row_values(1)
    if update_column not in headers:
        new_col_index = len(headers) + 1
        sheet.update_cell(1, new_col_index, update_column)
        update_col_index = new_col_index
        print(f"Added new column '{update_column}' at position {new_col_index}")
    else:
        update_col_index = headers.index(update_column) + 1

    updates = []
    for row_index, value in zip(row_indexes, update_values):
        updates.append({
            'range': f'{gspread.utils.rowcol_to_a1(row_index, update_col_index)}',
            'values': [[value]]
        })

    sheet.batch_update(updates)
    print(f"Updated {len(updates)} cells in column '{update_column}'")


def run_with_timeout(func, args=(), kwargs={}, timeout=5):
    def wrapper(result_queue):
        result = func(*args, **kwargs)
        result_queue.put(result)

    result_queue = multiprocessing.Queue()
    process = multiprocessing.Process(target=wrapper, args=(result_queue,))
    
    process.start()
    process.join(timeout)
    
    if process.is_alive():
        process.terminate()
        return None, False
    
    return result_queue.get(), True


def parse_template_message(template_message):
    index_pairs = list()
    start_index = None
    end_index = None
    open = True
    for i, character in enumerate(template_message):
        if character == "$":
            if i < len(template_message)-1:
                if template_message[i+1]=='{':
                    open = True
                    start_index = i 
        if open == True and character == "}":
            open = False
            end_index = i
            index_pairs.append((start_index, end_index))
    return index_pairs


def synthesise_message(template_message, index_pairs, name, subsystem, date, interview_time):
    variables = {
        "name": name,
        "subsystem": subsystem,
        "date": date,
        "interview_time": interview_time
    }
    # print(index_pairs)
    message = ""
    message += template_message[:index_pairs[0][0]]
    message += variables[template_message[index_pairs[0][0]+2:index_pairs[0][1]]]
    # print(message)
    for i in range(1, len(index_pairs)):
        index_pair = index_pairs[i]
        message += template_message[index_pairs[i-1][1]+1:index_pair[0]]
        # print(message)
        message += variables[template_message[index_pair[0]+2:index_pair[1]]]
        # print(message)
    message += template_message[index_pairs[-1][1]+1:]
    # print(message)
    return message


def send_message(name, phone_number, message, phone_number_backup = None):
    phone_number = str(phone_number)
    if len(phone_number)==10:
        phone_number = "+91" + phone_number
    i = 0
    timeout_tries = 0
    while True:
        try:
            # pywhatkit.sendwhatmsg_instantly(str(phone_number), message, wait_time=20, tab_close=True)
            # messenger.send_direct_message(str(phone_number), message, False)
            result, status = run_with_timeout(messenger.send_direct_message, args=(str(phone_number),message,False), timeout=PARAMS["timeout"])
            if status == False:
                timeout_tries += 1
                print(f"Attempt {timeout_tries}: Message timed out")
                if timeout_tries == PARAMS["max_timeout_tries"]:
                    if phone_number_backup:
                        return send_message(name, phone_number_backup,message)
                    else:
                        return False
                time.sleep(PARAMS["timeout_attempt_interval"])
            else:
                return True
        except Exception as e:
            i+=1
            print(f"Failed to send message. Retrying...{i}")
            print(name, phone_number, message)
            print(str(e))
            print(traceback.format_exc())


def calculate_time(i, start_time, end_time, duration):
    start = datetime.datetime.strptime(start_time, "%H:%M")
    end = datetime.datetime.strptime(end_time, "%H:%M")
    total_available_minutes = (end - start).total_seconds() / 60
    time_per_interview = duration
    if i * time_per_interview + duration > total_available_minutes:
        return False
    interview_start_time = start + datetime.timedelta(minutes=i * time_per_interview)
    return interview_start_time.strftime("%H:%M")


def format_phone_number(phone_number, phone_number_backup = None):
    if phone_number.lower().strip() == "same" and phone_number_backup:
        return str(format_phone_number(phone_number_backup))
    # Extract only digits from the input string
    digits = re.findall(r'\d+', phone_number)
    digits_only = ''.join(digits)
    
    # Check if the number has a country code (i.e., if it starts with a digit and is longer than 10 digits)
    if len(digits_only) > 10 and not digits_only.startswith('0'):
        # Check if the number already starts with '+', indicating a valid country code is likely present
        if phone_number.strip().startswith('+'):
            return f"+{digits_only}"
        else:
            return f"+{digits_only}"
    
    # If the number does not include a country code, add +91
    formatted_number = f"+91{digits_only}"
    
    return formatted_number


interview_count = len(details)

print("First 5 rows of sheet:")
print(details.head())
print("Number of interviews to schedule:", interview_count)

permission_to_continue()

with open("message.txt") as f:
    template_message = f.read()
    print("Contents of 'message.txt':")
    print(template_message)
index_pairs = parse_template_message(template_message)

print("Example message:")
print(synthesise_message(template_message, index_pairs, "<name>", PARAMS["target_subsystem"], PARAMS["date"], PARAMS["start_time"]))

permission_to_continue()

browser = webdriver.Chrome(options=chrome_options())
handles = browser.window_handles
for _, handle in enumerate(handles):
    if handle != browser.current_window_handle:
        browser.switch_to.window(handle)
        browser.close()

messenger = WhatsApp(browser = browser)

if "Notified_"+PARAMS["target_subsystem"] not in details.keys():
    details["Notified_"+PARAMS["target_subsystem"]] = ['']*len(details)

# Send all messages
i = 0
batch_count = 0
selected_indexes = list()
selected_indexes_values = list()
phone_number = PARAMS['testing']['recipient_phone_number']
for index, row in details.iterrows():
    # print(row)
    name = row[PARAMS['columns']['name']]
    phone_number_backup = format_phone_number(row[PARAMS["columns"]["mobile_number"]])
    if PARAMS["testing"]["test_mode"] != True:
        phone_number = format_phone_number(row[PARAMS['columns']['whatsapp_number']], phone_number_backup=row[PARAMS['columns']['mobile_number']])
    preference_columns = [PARAMS["columns"]["preference1"], PARAMS["columns"]["preference2"]]
    subsystem = row[preference_columns[PARAMS["subsystem_preference"]-1]]
    
    if i%PARAMS["at_once"] == 0:
       interview_time = calculate_time(batch_count, PARAMS["start_time"], PARAMS["end_time"], PARAMS["duration"])
       batch_count += 1

    if interview_time == False:
        print("End time reached.")
        print(f"{i} interviews scheduled.")
        break
    
    print(f"Attempting to schedule: Interview {i+1}/{len(details)} (Row {index+1}) @ {interview_time} for {subsystem} [{name}({phone_number})]")
    message = synthesise_message(template_message, index_pairs, name, subsystem, PARAMS["date"], interview_time)
    message_status = send_message(name, phone_number, message, phone_number_backup=phone_number_backup)
    selected_indexes.append(row["id"])
    if message_status == True:
        i+=1
        print(f"Scheduled.")
        selected_indexes_values.append(f"Notified: {PARAMS['date']} {interview_time}")
        append_row_to_sheet(interview_sheet, row['id'], interview_time)
    else:
        interview_count -= 1
        print(f"Sending the message timed out. The whatsapp number may be invalid.")
        selected_indexes_values.append("Message timed out")
    
    message_interval = PARAMS["message_interval"]
    if i <= 1 and message_interval < 15:
        message_interval += 10
    time.sleep(message_interval)
    print()

if PARAMS['testing']['test_mode'] != True:
    update_sheet_values(PARAMS["target_sheet_url"], PARAMS["target_worksheet_name"], "Notified_"+PARAMS["target_subsystem"], selected_indexes, selected_indexes_values)