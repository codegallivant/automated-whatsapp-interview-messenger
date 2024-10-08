# automated-whatsapp-interview-messenger
Automatic whatsapp messenger & scheduler for interviews at project MANAS

## Automated Setup

Run the following script through an Ubuntu shell or WSL instance (required for Windows machines).

```sh
wget https://raw.githubusercontent.com/codegallivant/automated-whatsapp-interview-messenger/master/setup.sh
chmod +x setup.sh
./setup.sh
```

> [!IMPORTANT]
> Modify the `settings.yaml` file in the placeholder sections.
> Download the `credentials.json` file and put it into the project root directory.

Once these steps are complete, run the `main.py` script.

```sh
python3 main.py
```

---

## Manual Setup

### Dependencies 
When building from source, install dependencies as follows -
1. Install pip modules - 
    ```bash
    pip install -r requirements.txt
    ```
2. Install a stable chrome and chrome driver version from [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/). (I used version 128). This is required as automated scripts can only work with chrome browsers for testing.

### Download service account credentials file
Set the service account credentials in ``sac_creds/credentials.json``. This is used to authenticate with the service account used to pull details from the google sheet and add notification status.

### Set template message
Modify/create a file called `message.txt` in the same directory as the script and set the contents. Use the format ``${variable_name}`` to insert variables like ``name`` or ``interview_time`` etc. For example:
```
Hello ${name}, this message is a confirmation of your Project Manas ${subsystem} interview. 
Your interview has been scheduled on ${date} at ${interview_time}. You are requested to come to the workshop 10 minutes prior to the given time slot. 
Our workshop location is: https://g.page/project-manas-manipal?share 
Reply to confirm.
```
Variables you can use in the message are ``interview_time``, ``name``, ``date``, ``subsystem``.

### Set settings
In `settings.yaml`, you can customize settings like so - 
```yaml
chrome_settings:
  path_to_chrome: "../chrome-linux64/chrome" # path to chrome
  path_to_chromedriver: "../chromedriver-linux64/chromedriver" # path to chrome driver
  headless: false # dont change from false, doesnt work

sheet_settings:
  # Source sheet details (Sheet being updated directly by google form)
  source_sheet_url: "<Insert sheet link here>" # Link to sheet (discard everything after the id i.e. from '/edit')
  source_worksheet_name: "Form Responses 1" # Sheet name
  # Target sheet details (Sheet being modified by program and synced with source sheet)
  target_sheet_url: "<Insert sheet link here>" # Link to sheet (discard everything after the id i.e. from '/edit')
  target_worksheet_name: "Form Responses 1" # Sheet name

  # Interview score sheet details
  interview_score_url: "<Insert sheet link here>" # Link to sheet (discard everything after the id i.e. from '/edit')


interview_time_settings:
  date: "27/08/2024" # Date(DD/MM/YYYY) of interviews to be scheduled.
  start_time: "10:00" # Time(hh:mm 24hr) the first interview should start at
  end_time: "21:00" # Maximum time(hh:mm 24hr) by which the last interview should end
  duration: 25 # Duration in minutes of each interview
  at_once: 6 # Number of interviews to schedule for the same time 

subsystem_settings:
  subsystem_preference: 1 # Subsystem preference number to be chosen for the interview.
  target_subsystem: "Artificial Intelligence" # Subsystem to schedule interviews for. Set empty string("") or null for no restrictions. This must be the same string as in the sheet

message_settings:
  message_interval: 10 # Number of seconds to wait between sending consecutive whatsapp messages
  timeout: 20 # Timeout for sending message (required to handle invalid phone numbers)
  max_timeout_tries: 3 # Maximum number of tries to try sending timing out messages
  timeout_attempt_interval: 5

notifier_settings:
  notifier: "<your-name>" # Chrome user data will be attached to this name. 
  check_notifier_column: false # If set to true, only rows where member notifier matches the notifier specified will be picked 

columns:
  name: "Full Name"
  registration_number: "Registration No. " 
  branch: "Branch"
  whatsapp_number: "WhatsApp Number"
  mobile_number: "Mobile Number"
  email_id: "Email Address"
  learner_id: "Learner ID (@learner.manipal.edu)"
  preference1: "First Preference of Subsystem"
  preference2: "Second Preference of Subsystem"
  description: "Tell us about yourself."
  first_year: "This form is only for First Year Students. Are you in First Year?"
  notifier: "MemberNotifier"

testing: # for testing without sending messages to others
  test_mode: false
  recipient_phone_number: "<phone number>"
  send_message: true
  update_score_sheet: true
  update_target_sheet: true
```

### Run
1. Install a browser and driver from [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/) and other dependencies
2. Download the service account credentials file into `credentials.json`
3. Set template message in `message.txt`
4. Set settings in `settings.yaml`
5. Run the file - 
    ```bash
    python3 main.py
    ```

### Sync target with source sheet
Run 
```bash
python3 sync_sheets.py
```
to sync the target sheet with the source sheet. Avoid running this too frequently.
