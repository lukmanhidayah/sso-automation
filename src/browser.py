import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service

def setup_driver(headless=True):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    chrome_options.add_argument('--disable-gpu')
    chrome_options.add_argument('--disable-blink-features=AutomationControlled')  # Hindari deteksi bot
    chrome_options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36')

    # Reduce noisy Chrome/Chromedriver logs and disable push/translate features
    chrome_options.add_argument('--log-level=3')              # Only fatal
    chrome_options.add_argument('--disable-logging')
    chrome_options.add_argument('--disable-notifications')
    chrome_options.add_argument('--disable-features=PushMessaging,Translate,MediaRouter')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging', 'enable-automation'])
    chrome_options.add_experimental_option('useAutomationExtension', False)

    # Respect container-provided Chrome/Chromedriver paths if present
    chrome_bin = os.environ.get('CHROME_BIN')
    if chrome_bin:
        chrome_options.binary_location = chrome_bin

    chromedriver_path = os.environ.get('CHROMEDRIVER')
    if chromedriver_path:
        # Redirect Chromedriver service logs to OS null device
        service = Service(executable_path=chromedriver_path, log_output=os.devnull)
        return webdriver.Chrome(service=service, options=chrome_options)
    else:
        # Default service also with quiet logging
        service = Service(log_output=os.devnull)
        return webdriver.Chrome(service=service, options=chrome_options)
