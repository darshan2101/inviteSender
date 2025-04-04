#!/usr/bin/env python3

import subprocess
import sys
import os
import time
import io
import shutil
import urllib.request
import argparse
import unicodedata
import pyperclip



# STEP 0: INSTALL PYTHON DEPENDENCIES IF MISSING
required_packages = ['pandas', 'selenium', 'PyPDF2', 'reportlab', 'webdriver-manager']

def install_packages():
    subprocess.check_call([sys.executable, "-m", "pip", "install", *required_packages])

try:
    import pandas as pd
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from PyPDF2 import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains
except ImportError:
    print("📦 Installing missing packages...")
    install_packages()
    import pandas as pd
    from selenium import webdriver
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.chrome.options import Options
    from PyPDF2 import PdfReader, PdfWriter
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.common.action_chains import ActionChains

def is_gujarati(text):
    return any("GUJARATI" in unicodedata.name(char, "") for char in text)

def wait_for_element(xpath, timeout=15):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.XPATH, xpath))
    )

# STEP 1: CHECK CHROME INSTALL
def is_chrome_installed():
    return shutil.which("google-chrome") is not None

def install_chrome():
    print("🌐 Installing Google Chrome (Ubuntu)...")
    subprocess.run(["wget", "https://dl.google.com/linux/direct/google-chrome-stable_current_amd64.deb"])
    subprocess.run(["sudo", "apt", "install", "./google-chrome-stable_current_amd64.deb", "-y"])
    subprocess.run(["rm", "google-chrome-stable_current_amd64.deb"])

if not is_chrome_installed():
    install_chrome()

# STEP 2: CHECK & DOWNLOAD GUJARATI FONT
# STEP 2: USE LOCAL FONTS (NO DOWNLOAD NEEDED)
GUJARATI_FONT_PATH = "./Noto_Sans_Gujarati/static/NotoSansGujarati-Regular.ttf"
ENGLISH_FONT_PATH = "./Roboto/static/Roboto-Regular.ttf"

# Check if font files exist
if not os.path.exists(GUJARATI_FONT_PATH):
    print(f"❌ Gujarati font not found at {GUJARATI_FONT_PATH}")
    sys.exit(1)

if not os.path.exists(ENGLISH_FONT_PATH):
    print(f"❌ English font not found at {ENGLISH_FONT_PATH}")
    sys.exit(1)

# Register both fonts (currently using GujaratiFont for everything)
pdfmetrics.registerFont(TTFont('GujaratiFont', GUJARATI_FONT_PATH))
pdfmetrics.registerFont(TTFont('EnglishFont', ENGLISH_FONT_PATH))


# ==== CONFIG ====
TEMPLATE_PDF = "invitation_template.pdf"
CSV_PATH = "contacts.csv"
OUTPUT_FOLDER = "output"
WAIT_TIME = 10  # seconds to wait for WhatsApp Web to load

parser = argparse.ArgumentParser()
parser.add_argument("--limit", type=int, default=5, help="Number of contacts to process (default: 5)")
args = parser.parse_args()
NUM_CONTACTS_TO_SEND = args.limit

MESSAGE_TEMPLATE = "Hey {name}, here's your personalized invitation . Please see the attached PDF."

# Create output dir
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ==== LOAD CONTACTS ====
contacts = pd.read_csv(CSV_PATH).head(NUM_CONTACTS_TO_SEND)

# ==== SETUP SELENIUM CHROME ====
options = Options()
options.binary_location = shutil.which("google-chrome")
options.add_argument("--user-data-dir=./User_Data")  # saves login session
options.add_argument("--profile-directory=Default")
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Go to WhatsApp Web
driver.get("https://web.whatsapp.com")

print("⌛ Waiting for WhatsApp Web login...")
time.sleep(WAIT_TIME)

# ==== FUNCTIONS ====

def create_custom_pdf(name, output_path):
    packet = io.BytesIO()
    can = canvas.Canvas(packet, pagesize=letter)

    font_name = "GujaratiFont" if is_gujarati(name) else "EnglishFont"
    can.setFont(font_name, 16)
    can.drawString(200, 200, name)
    can.save()

    packet.seek(0)
    overlay = PdfReader(packet)
    template = PdfReader(open(TEMPLATE_PDF, "rb"))

    output = PdfWriter()
    page = template.pages[0]
    page.merge_page(overlay.pages[0])
    output.add_page(page)

    with open(output_path, "wb") as out:
        output.write(out)
def send_message_and_file(phone, name, file_path):
    url = f"https://web.whatsapp.com/send?phone={phone}&text&app_absent=0"
    driver.get(url)

    # Wait for message box to appear
    wait_for_element('//div[@aria-label="Type a message" and @contenteditable="true"]')

    try:
        # Prepare message
        message = MESSAGE_TEMPLATE.format(name=name)
        pyperclip.copy(message)

        # Wait for message box to be clickable
        msg_box_xpath = '//div[@aria-label="Type a message" and @contenteditable="true"]'
        msg_box = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, msg_box_xpath))
        )

        # Scroll to the element to avoid interception
        driver.execute_script("arguments[0].scrollIntoView(true);", msg_box)
        time.sleep(0.5)  # Let any animations finish

        # Use ActionChains to paste the message
        actions = ActionChains(driver)
        actions.move_to_element(msg_box).click().key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        time.sleep(0.5)

        # Press Enter to send
        msg_box.send_keys(Keys.ENTER)
        time.sleep(2)

        # Attach file
        attach_btn = driver.find_element(By.XPATH, '//span[@data-icon="clip"]')
        attach_btn.click()
        time.sleep(1)

        file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
        file_input.send_keys(os.path.abspath(file_path))
        time.sleep(2)

        # Click send button
        send_btn = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_btn.click()
        time.sleep(3)

        print(f"✅ Sent to {name} ({phone})")

    except Exception as e:
        print(f"❌ Failed for {name} ({phone}): {e}")
        driver.save_screenshot(f"screenshot_failure_{name.replace(' ', '_')}.png")
        print(f"📸 Screenshot saved for {name}")

# ==== PROCESS CONTACTS ====
for index, row in contacts.iterrows():
    name = str(row['Name']).strip()
    phone = str(row['PhoneNumber']).strip()

    pdf_path = os.path.join(OUTPUT_FOLDER, f"invitation_{name.replace(' ', '_')}.pdf")
    create_custom_pdf(name, pdf_path)
    send_message_and_file(phone, name, pdf_path)

print("🎉 All done! PDFs personalized and messages sent.")
