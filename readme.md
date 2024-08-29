# automated-whatsapp-interview-messenger
Automatic whatsapp messenger & scheduler for interviews at project MANAS

## Dependencies
1. ``pip install -r requirements.txt``
2. Install a stable chrome and chrome driver version from [Chrome for Testing](https://googlechromelabs.github.io/chrome-for-testing/). (I used version 128). This is required as automated scripts can only work with chrome browsers in testing.
3. Modify the python ``alright`` package as follows - 
    \
    In ``site-packages/alright/__init__.py``, replace the indented block under ``if not browser:`` with:
    ```python
    self.chrome_options.binary_location = "path/to/chrome"
    self.chrome_options.executable_path="/path/to/chromedriver"
    browser = webdriver.Chrome(options=self.chrome_options)
    
    handles = browser.window_handles
    for _, handle in enumerate(handles):
        if handle != browser.current_window_handle:
            browser.switch_to.window(handle)
            browser.close()
    ```

## Set template message
Modify/create a file called "message.txt" in the same directory as the script and set the contents. Use ${variable_name} for variables like ``name`` or ``interview_time`` etc. For example:
```
Hello ${name}. Greetings from Project MANAS. Your interview for ${subsystem} is scheduled for ${date} at ${interview_time}.
```

## Run
1. Add interviewee details in a file called ``details.csv`` in the same directory
2. ```bash
python3 main.py
```