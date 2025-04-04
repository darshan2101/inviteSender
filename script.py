#!/usr/bin/env python3

import subprocess
import sys
import os
import time
import io
import shutil
import urllib.request
import argparse


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
FONT_FILE = "NotoSansGujarati-Regular.ttf"
FONT_URL = "https://github.com/google/fonts/raw/main/ofl/notosansgujarati/NotoSansGujarati-Regular.ttf"

if not os.path.exists(FONT_FILE):
    print("🔤 Downloading Gujarati-compatible font (NotoSansGujarati)...")
    urllib.request.urlretrieve(FONT_URL, FONT_FILE)

pdfmetrics.registerFont(TTFont('GujaratiFont', FONT_FILE))

# ==== CONFIG ====
TEMPLATE_PDF = "invitation_template.pdf"
CSV_PATH = "contacts.csv"
OUTPUT_FOLDER = "output"
WAIT_TIME = 10  # seconds to wait for WhatsApp Web to load

parser = argparse.ArgumentParser()
parser.add_argument("--limit", type=int, default=5, help="Number of contacts to process (default: 5)")
args = parser.parse_args()
NUM_CONTACTS_TO_SEND = args.limit

MESSAGE_TEMPLATE = "Hey {name}, here's your personalized invitation 🎉. Please see the attached PDF."

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
    can.setFont("GujaratiFont", 16)
    can.drawString(200, 500, name)
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
    time.sleep(8)

    try:
        msg_box = driver.find_element(By.XPATH, '//div[@title="Type a message"]')
        msg_box.click()
        msg_box.send_keys(MESSAGE_TEMPLATE.format(name=name))
        msg_box.send_keys(Keys.ENTER)
        time.sleep(2)

        attach_btn = driver.find_element(By.XPATH, '//span[@data-icon="clip"]')
        attach_btn.click()
        time.sleep(1)

        file_input = driver.find_element(By.CSS_SELECTOR, 'input[type="file"]')
        file_input.send_keys(os.path.abspath(file_path))
        time.sleep(2)

        send_btn = driver.find_element(By.XPATH, '//span[@data-icon="send"]')
        send_btn.click()
        time.sleep(3)

        print(f"✅ Sent to {name} ({phone})")

    except Exception as e:
        print(f"❌ Failed for {name} ({phone}): {e}")

# ==== PROCESS CONTACTS ====
for index, row in contacts.iterrows():
    name = str(row['Name']).strip()
    phone = str(row['PhoneNumber']).strip()

    pdf_path = os.path.join(OUTPUT_FOLDER, f"invitation_{name.replace(' ', '_')}.pdf")
    create_custom_pdf(name, pdf_path)
    send_message_and_file(phone, name, pdf_path)

print("🎉 All done! PDFs personalized and messages sent.")
