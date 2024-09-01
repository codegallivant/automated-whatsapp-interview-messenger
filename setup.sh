echo "Working in Current Directory"

mkdir -p ManasInterviews2024 && cd ManasInterviews2024

# Chrome Browser Executable
wget https://storage.googleapis.com/chrome-for-testing-public/128.0.6613.86/linux64/chrome-linux64.zip
unzip chrome-linux64.zip
echo "Downloaded Chrome Browser Binary"

# Chrome Driver
wget https://storage.googleapis.com/chrome-for-testing-public/128.0.6613.86/linux64/chromedriver-linux64.zip
unzip chromedriver-linux64.zip
echo "Downloaded Chrome Driver"

# Our Fancy Jazzy Epic Code
git clone https://github.com/codegallivant/automated-whatsapp-interview-messenger.git
echo "Downloaded Automation Script"
cd automated-whatsapp-interview-messenger

# Requirements
echo "Installing Python Requirements"
pip install -r requirements.txt

echo "Done. Run 'python main.py' to start the automation script."