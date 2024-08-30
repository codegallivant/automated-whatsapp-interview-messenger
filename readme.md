# automated-whatsapp-interview-messenger
Automatic whatsapp messenger & scheduler for interviews at project MANAS

## Dependencies 
When building from source, install dependencies as follows -
1. Install pip modules - 
    ```bash
    pip install -r requirements.txt
    ```
2. Install a stable chrome and chrome driver version from [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/). (I used version 128). This is required as automated scripts can only work with chrome browsers for testing.
3. Modify the python `alright` package as follows - 
    \
    In `site-packages/alright/__init__.py`, replace the following block of code under ``if not browser:``:
    ```python    
    browser = webdriver.Chrome(
    ChromeDriverManager().install(),
    options=self.chrome_options,
    )
    ```
    with the following:
    ```python
    self.chrome_options.binary_location = "path/to/chrome"
    self.chrome_options.executable_path="path/to/chromedriver"
    browser = webdriver.Chrome(options=self.chrome_options)
    ```

## Set template message
Modify/create a file called `message.txt` in the same directory as the script and set the contents. Use the format ``${variable_name}`` to insert variables like ``name`` or ``interview_time`` etc. For example:
```
Hello ${name}. Greetings from Project MANAS. Your interview for ${subsystem} is scheduled for ${date} at ${interview_time}.
```
Variables you can use in the message are ``interview_time``, ``name``, ``date``, ``subsystem``.

## Set parameters
In `parameters.yaml`, you can customize parameters like so - 
```yaml
date: "27/08/2024" # Date of interviews to be scheduled. Can be in any format
start_time: "10:00" # Time(same format, 24hr) the first interview should start at
end_time: "21:00" # Maximum time(same format, 24hr) at which the last interview should end by
padding_minutes: 15 # Number of minutes to wait between each interview
duration: 25 # Duration in minutes of each interview
at_once: 6 # Number of interviews to schedule for the same time 

subsystem_preference: 1 # Subsystem preference number to be chosen for the interview.
target_subsystem: "Artificial Intelligence" # Subsystem to schedule interviews for. Set empty string("") or null for no restrictions.

message_interval: 10 # Number of seconds to wait between sending consecutive whatsapp messages

timeout: 20 # Timeout for sending message (required to handle invalid phone numbers)
max_timeout_tries: 3 # Maximum number of tries to try sending timing out messages

columns: # Enter names of columns in the CSV file
  name: "Full Name"
  whatsapp_number: "WhatsApp Number"
  preference1: "First Preference of Subsystem"
  preference2: "Second Preference of Subsystem"
  first_year: "This form is only for First Year Students. Are you in First Year?"
```

## Run
1. Add interviewee details in a file called `details.csv` in the same directory
2. Run the file - 
    ```bash
    python3 main.py
    ```