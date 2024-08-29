# import pywhatkit
import pandas as pd
import yaml
import traceback
import datetime
import time
from alright import WhatsApp

with open("parameters.yaml") as stream:
    try:
        PARAMS = yaml.safe_load(stream)
        print(PARAMS)
    except yaml.YAMLError as exc:
        print(exc)


details = pd.read_csv("details.csv")


def synthesise_message(name, subsystem, date, interview_time):
    message_parts = [
"Hello ",
name,
". Greetings from Project MANAS. Your interview for ",
subsystem,
" is scheduled for ",
date,
" at ",
interview_time,
"."
]
    return "".join(message_parts)


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
    new_hours = this_start_time_obj.hour
    new_minutes = this_start_time_obj.minute

    print(hours, minutes)
    this_end_datetime_obj = this_start_datetime_obj + datetime.timedelta(minutes=duration)
    this_end_time_obj = this_end_datetime_obj.time()
    this_end_hour = this_end_time_obj.hour
    this_end_minutes = this_end_time_obj.minute
    # minutes_increase = minutes + i*(duration+padding_minutes)
    # hours_increase = minutes//60
    # new_minutes = minutes_increase % 60
    # new_hours = hours+hours_increase
    # new_time = t + datetime.timedelta(minutes = minutes_increase)
    # new_hours = new_time.hours
    # new_minutes = new_time.minutes
    end_hour = int(end_time.split(":")[0])
    end_minutes = int(end_time.split(":")[1])
    if this_end_hour == end_hour:
        if this_end_minutes >= end_minutes:
            return -1
    if this_end_hour > end_hour:
        return -1
    if new_minutes < 10:
        new_minutes = "0"+str(new_minutes)
    interview_time = str(new_hours)+":"+str(new_minutes)
    print(interview_time)
    return interview_time


# Send all messages

i = 0
for index, row in details.iterrows():
    print(row)
    name = row["Full Name"]
    phone_number = row["WhatsApp Number"]
    registration_numbers = row["Registration No. "]
    preference_columns = ["First Preference of Subsystem", "Second Preference of Subsystem"]
    subsystem = row[preference_columns[PARAMS["subsystem_preference"]-1]]
    first_year = row["This form is only for First Year Students. Are you in First Year?"]
    if first_year != "Yes":
        continue
    interview_time = calculate_time(i, PARAMS["start_time"], PARAMS["end_time"], PARAMS["duration"], PARAMS["padding_minutes"])
    if interview_time == -1:
        print("End time reached.")
        print(f"{i} interviews scheduled.")
        break
    message = synthesise_message(name, subsystem, PARAMS["date"], interview_time)
    send_message(name, phone_number, message)
    time.sleep(PARAMS["message_interval"])
    i+=1