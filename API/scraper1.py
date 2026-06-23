# coding=utf-8
import os
import re
import time
import warnings
import sys
import cv2
import pytesseract
from PIL import Image
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import xlwt

# ------------------ Tesseract Path ------------------
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
# -----------------------------------------------------

if not sys.warnoptions:
    warnings.simplefilter("ignore")

# ------------------ User Inputs ------------------
college = input("Enter the college code\n").upper()
year = input('Enter the year\n')
branch = input('Please enter the branch\n').upper()
low = int(input('Enter starting USN\n'))
high = int(input('Enter last USN\n')) + 1
semc = input('Enter the Semester\n')

dip = 'Y' if low >= 400 else 'N'

subcode = 52
iloop = 8
if semc in ['3', '4']:
    iloop = 9
    subcode = 58
    if dip == 'Y':
        iloop = 10
        subcode = 64

# ------------------ Setup Chrome Driver ------------------
service = Service(ChromeDriverManager().install())
options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=service, options=options)

url = "https://results.vtu.ac.in/JJEcbcs25/index.php"

# ------------------ Helper Functions ------------------
def clean_captcha(text):
    """Clean OCR output for CAPTCHA"""
    return re.sub(r'[^A-Za-z0-9]', '', text.strip())

def fetch_result(usn):
    """Fetch result for a single USN with CAPTCHA OCR and manual fallback"""
    retry_count = 0
    max_auto_attempts = 3  # Automatic OCR attempts

    while True:
        try:
            driver.get(url)
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.NAME, "lns"))
            )

            # Enter USN
            usn_input = driver.find_element(By.NAME, "lns")
            usn_input.clear()
            usn_input.send_keys(usn)

            # Screenshot CAPTCHA
            captcha_img = driver.find_element(By.XPATH, "//img[@alt='CAPTCHA code']")
            captcha_img.screenshot("cap.png")

            # Preprocess for OCR
            img_cv = cv2.imread("cap.png")
            gray = cv2.cvtColor(img_cv, cv2.COLOR_BGR2GRAY)
            gray = cv2.GaussianBlur(gray, (3, 3), 0)
            _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY_INV)
            cv2.imwrite("cap_processed.png", thresh)
            img = Image.open("cap_processed.png")

            # OCR
            captcha_text = pytesseract.image_to_string(
                img,
                config='--psm 8 -c tessedit_char_whitelist=ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789'
            )
            captcha_text = clean_captcha(captcha_text)

            # If OCR failed after N attempts → ask manual
            if len(captcha_text) == 0 or retry_count >= max_auto_attempts:
                captcha_text = input(f"[Manual Entry] Enter CAPTCHA for USN {usn}: ").strip()
                retry_count = 0  # reset after manual entry

            print(f"Using CAPTCHA: {captcha_text}")

            cap_input = driver.find_element(By.NAME, "captchacode")
            cap_input.clear()
            cap_input.send_keys(captcha_text)

            # Submit
            driver.find_element(By.ID, "submit").click()
            time.sleep(2)

            # Handle alert (invalid captcha / invalid USN)
            try:
                alert = driver.switch_to.alert
                alert_text = alert.text
                alert.accept()

                if "Invalid captcha" in alert_text:
                    print(f"❌ Invalid CAPTCHA for {usn}. Retrying (manual may be needed)...")
                    retry_count += 1
                    continue
                elif "University Seat Number" in alert_text:
                    print(f"No results for USN: {usn}")
                    return None
            except:
                pass

            # If reached here → parse result
            soup = BeautifulSoup(driver.page_source, "html.parser")
            tds = soup.find_all('td')
            divs = soup.find_all('div', attrs={'class': 'col-md-12'})
            divCell = soup.find_all('div', attrs={'class': 'divTableCell'})

            if len(tds) == 0 or len(divs) < 6:
                print(f"Invalid USN/Data for {usn}")
                return None

            sem = divs[5].div.text.strip('Semester : ')
            if sem != semc:
                print(f"Semester mismatch for {usn}")
                return None

            # Build record
            record = f"{re.sub('[!@#$:]', '', tds[1].text)},{re.sub('[!@#$:]', '', tds[3].text)},"

            sortList1 = []
            for i in range(6, subcode, 6):
                code = divCell[i].text[-3:] if divCell[i].text[-3:].isdigit() else divCell[i].text[-2:]
                sortList1.append(code)
            sortList1.sort()

            ilist = []
            for i in range(iloop):
                for j in range(6, subcode, 6):
                    code = divCell[j].text[-3:] if divCell[j].text[-3:].isdigit() else divCell[j].text[-2:]
                    if code == sortList1[i] and j not in ilist:
                        ilist.append(j)

            for l in ilist:
                for j in range(l, l + 6):
                    if j == l + 1:
                        continue
                    char = divCell[j].text.strip()
                    record += f"{int(char)}," if char.isdigit() else f"{char},"

            return record.strip(',')

        except Exception as e:
            print(f"⚠️ Error fetching USN {usn}: {e}")
            retry_count += 1
            time.sleep(2)

# ------------------ Main Loop ------------------
results = []
for u in range(low, high):
    usn = f'{college}{year}{branch}{str(u).zfill(3)}'
    print(f"\nFetching result for {usn}")
    res = fetch_result(usn)
    if res:
        results.append(res)

# ------------------ Write Excel ------------------
if results:
    book = xlwt.Workbook()
    ws = book.add_sheet('Sheet1')
    style = xlwt.XFStyle()

    for i, line in enumerate(results):
        row = line.split(',')
        for j, val in enumerate(row):
            ws.write(i, j, int(val) if val.isdigit() else val, style)

    os.makedirs('ExcelFiles', exist_ok=True)
    filename = f"{college}{year}{branch}{low}-{high-1}{'DIP' if dip=='Y' else ''}.xls"
    book.save(os.path.join('ExcelFiles', filename))
    print(f"\n✅ Results saved to Excel: ExcelFiles/{filename}")
else:
    print("\n⚠️ No results fetched. Please check inputs and CAPTCHA handling.")

# ------------------ Cleanup ------------------
for file in ['cap.png', 'cap_processed.png']:
    try:
        os.remove(file)
    except:
        pass

driver.quit()
