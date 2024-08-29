import pandas as pd
import yaml
import traceback
import datetime
import time
from alright import WhatsApp
import sys


def permission_to_continue():
    while True:
        print()
        response = input("Continue execution? (y/n) ")
        print()
        if response in ["y","Y"]:
            return
        elif response in ["n","N"]:
            sys.exit()    


def print_params(params):
    print()
    print("Parameters:")
    for param in params.keys():
        if isinstance(params[param], dict):
            print()
            print(param)
            print_params(params[param])
        else:
            print(param, ":", params[param])


with open("parameters.yaml") as stream:
    try:
        PARAMS = yaml.safe_load(stream)
        print_params(PARAMS)
        print()
    except yaml.YAMLError as exc:
        print(exc)

permission_to_continue()

details = pd.read_csv("details.csv")

print("First 5 rows of 'details.csv':")
print(details.head())
permission_to_continue()


def parse_template_message():
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

with open("message.txt") as f:
    template_message = f.read()
    print("Contents of 'message.txt':")
    print(template_message)
index_pairs = parse_template_message()


def synthesise_message(name, subsystem, date, interview_time):
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

print("Example message:")
print(synthesise_message("Abc Dfg", "Artificial Intelligence", "27/08/2024", "19:00"))

permission_to_continue()

messenger = WhatsApp()
def send_message(name, phone_number, message):
    phone_number = str(phone_number)
    if len(phone_number)==10:
        phone_number = "+91" + phone_number
    i = 0
    while True:
        try:
            # pywhatkit.sendwhatmsg_instantly(str(phone_number), message, wait_time=20, tab_close=True)
            messenger.send_direct_message(str(phone_number), message, False)
            break
        except Exception as e:
            i+=1
            print(f"Failed to send message. Retrying...{i}")
            print(name, phone_number, message)
            print(str(e))
            print(traceback.format_exc())


def calculate_time(i, start_time, end_time, duration, padding_minutes):
    hours = int(start_time.split(":")[0])
    minutes = int(start_time.split(":")[1])
    
    minutes_increase = i*(duration+padding_minutes)
    
    this_start_datetime_obj = (datetime.datetime.combine(datetime.date.today(), datetime.time(hours, minutes))+datetime.timedelta(minutes=minutes_increase))
    this_start_time_obj = this_start_datetime_obj.time()
    this_start_time_hours = this_start_time_obj.hour
    this_start_time_minutes = this_start_time_obj.minute

    this_end_datetime_obj = this_start_datetime_obj + datetime.timedelta(minutes=duration)
    this_end_time_obj = this_end_datetime_obj.time()
    this_end_time_hour = this_end_time_obj.hour
    this_end_time_minutes = this_end_time_obj.minute
    end_hour = int(end_time.split(":")[0])
    end_minutes = int(end_time.split(":")[1])
    
    if this_end_time_hour == end_hour:
        if this_end_time_minutes >= end_minutes:
            return False
    if this_end_time_hour > end_hour:
        return False
    
    if this_start_time_minutes < 10:
        this_start_time_minutes = "0"+str(this_start_time_minutes)
    
    interview_time = str(this_start_time_hours)+":"+str(this_start_time_minutes)
    # print(interview_time)
    
    return interview_time


def validate_row(first_year, subsystem):
    return (first_year == "Yes") and (subsystem == PARAMS["target_subsystem"] or PARAMS["target_subsystem"] in [None, ""])


# Send all messages
i = 0
batch_count = 0
for index, row in details.iterrows():
    # print(row)
    name = row[PARAMS['columns']['name']]
    phone_number = row[PARAMS['columns']['whatsapp_number']]
    preference_columns = [PARAMS["columns"]["preference1"], PARAMS["columns"]["preference2"]]
    subsystem = row[preference_columns[PARAMS["subsystem_preference"]-1]]
    first_year = row[PARAMS["columns"]["first_year"]]

    if not validate_row(first_year, subsystem):
        print("Skipping..")
        continue
    
    if i%PARAMS["at_once"] == 0:
       interview_time = calculate_time(batch_count, PARAMS["start_time"], PARAMS["end_time"], PARAMS["duration"], PARAMS["padding_minutes"])
       batch_count += 1
    if interview_time == False:
        print("End time reached.")
        print(f"{i} interviews scheduled.")
        break
    message = synthesise_message(name, subsystem, PARAMS["date"], interview_time)
    send_message(name, phone_number, message)
    print(f"{index}: Interview", "for", subsystem,"scheduled at", interview_time+".", f"Message sent to {name}(Whatsapp number: {str(phone_number)}).")
    i+=1
    time.sleep(PARAMS["message_interval"])