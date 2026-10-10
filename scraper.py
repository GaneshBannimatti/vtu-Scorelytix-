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
from selenium.common.exceptions import TimeoutException, WebDriverException

from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# SETTINGS
# ============================================================

MAX_RETRIES = None  # Retry forever until the result is fetched
PAGE_LOAD_TIMEOUT = 30
ELEMENT_TIMEOUT = 15
RETRY_DELAY = 3

warnings.filterwarnings("ignore")


# ============================================================
# GLOBALS
# ============================================================

url = ""
college = ""
year = ""
branch = ""
low = 0
high = 0
semc = ""


# ============================================================
# TESSERACT
# ============================================================

def configure_tesseract():
    candidates = [
        r"C:\Program Files\Tesseract-OCR\tesseract.exe",
        r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
    ]

    for path in candidates:
        if os.path.exists(path):
            pytesseract.pytesseract.tesseract_cmd = path
            print(f"✅ Tesseract found: {path}")
            return

    print("⚠️ Tesseract executable was not found in the default locations.")
    print("   If CAPTCHA OCR fails, update the Tesseract path in scraper.py.")


configure_tesseract()


# ============================================================
# GET INPUTS
# ============================================================

def get_inputs():
    if len(sys.argv) == 8:
        college_value = sys.argv[1].upper()
        year_value = sys.argv[2]
        branch_value = sys.argv[3].upper()
        low_value = int(sys.argv[4])
        high_value = int(sys.argv[5]) + 1
        sem_value = sys.argv[6]
        url_value = sys.argv[7]
    else:
        college_value = input("Enter the college code: ").strip().upper()
        year_value = input("Enter the year: ").strip()
        branch_value = input("Enter the branch: ").strip().upper()

        low_value = int(input("Enter starting USN: ").strip())
        high_value = int(input("Enter last USN: ").strip()) + 1

        sem_value = input("Enter the Semester: ").strip()
        url_value = input("Enter VTU Result URL: ").strip()

    match = re.search(r"\d+", str(sem_value))

    if not match:
        raise ValueError(
            "Invalid semester. Enter a value such as 3, 4, 5 or 6."
        )

    sem_value = match.group(0)

    if not 1 <= int(sem_value) <= 10:
        raise ValueError("Semester must be between 1 and 10.")

    print(f"\n🎓 Selected Semester: {sem_value}")

    exact_codes = get_semester_subject_codes(sem_value)

    if exact_codes:
        print("🎯 Exact subject codes:")
        print("   " + ", ".join(exact_codes))
    else:
        print(
            f"🎯 Fallback subject prefix: "
            f"{get_target_subject_prefix(sem_value)}"
        )

    return (
        college_value,
        year_value,
        branch_value,
        low_value,
        high_value,
        sem_value,
        url_value,
    )


# ============================================================
# EXACT SEMESTER SUBJECT CONFIGURATION
# ============================================================

SEMESTER_SUBJECT_CODES = {
    "6": [
        "BCS601",
        "BCS602",
        "BCS685",
        "BCSL606",
        "BNSK658",
        "BIKS609",
        "BEC654A",
        "BCS613D",
        "BISL657E",
    ]
}


def get_target_subject_prefix(semester):
    match = re.search(r"\d+", str(semester))

    if not match:
        raise ValueError(f"Invalid semester: {semester}")

    return f"BCS{match.group(0)}".upper()


def get_semester_subject_codes(semester):
    # Do not hard-code a fixed list of subject codes. VTU subject codes
    # can vary by scheme, electives, and course type. The result page is
    # already for the selected semester, so parse every valid subject row.
    match = re.search(r"\d+", str(semester))
    if not match:
        raise ValueError(f"Invalid semester: {semester}")
    return None


def is_subject_for_semester(code, semester):
    """Keep only subjects whose code indicates the requested semester.

    Supports common VTU formats such as BCS601 / BNSK658 and older
    scheme formats such as 21CS61. Backlog subjects from other semesters
    are excluded even when VTU displays them on the same result page.
    """
    code = str(code or "").strip().upper()
    match = re.search(r"\d+", str(semester))
    if not code or code in {"SUBJECT CODE", "CODE"} or not match:
        return False

    requested_semester = match.group(0)
    if not re.fullmatch(r"[A-Z0-9]+", code):
        return False

    # Scheme/year-prefixed codes, e.g. 21CS61 or 22BCS601:
    # take the first digit after the alphabetic branch segment.
    year_prefixed = re.match(r"^\d{2,}[A-Z]+(\d+)", code)
    if year_prefixed:
        semester_digit = year_prefixed.group(1)[0]
        return semester_digit == requested_semester

    # Usual VTU codes, e.g. BCS601, BCSL606, BNSK658, BIKS609,
    # BEC654A, BISL657E: semester is the first digit after the prefix.
    standard = re.match(r"^[A-Z]+(\d+)", code)
    if standard:
        semester_digit = standard.group(1)[0]
        return semester_digit == requested_semester

    return False


# ============================================================
# SETUP SELENIUM
# ============================================================

def create_driver():
    print("\n🌐 Starting Chrome...")

    service = Service(
        ChromeDriverManager().install()
    )

    options = webdriver.ChromeOptions()

    options.add_argument(
        "--disable-blink-features=AutomationControlled"
    )
    options.add_argument("--start-maximized")

    # IMPORTANT:
    # Headless mode is disabled so you can see Chrome while debugging.
    # Add --headless again later if everything works.
    # options.add_argument("--headless=new")

    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--ignore-certificate-errors")
    options.add_argument("--allow-insecure-localhost")

    options.add_experimental_option(
        "excludeSwitches",
        ["enable-automation"]
    )

    options.add_experimental_option(
        "useAutomationExtension",
        False
    )

    browser = webdriver.Chrome(
        service=service,
        options=options
    )

    browser.set_page_load_timeout(
        PAGE_LOAD_TIMEOUT
    )

    print("✅ Chrome started successfully.")

    return browser


driver = None


# ============================================================
# CAPTCHA
# ============================================================

def clean_captcha(text):
    return re.sub(
        r"[^A-Za-z0-9]",
        "",
        str(text).strip()
    )


def solve_captcha(image_path):
    img_cv = cv2.imread(image_path)

    if img_cv is None:
        return ""

    gray = cv2.cvtColor(
        img_cv,
        cv2.COLOR_BGR2GRAY
    )

    gray = cv2.GaussianBlur(
        gray,
        (3, 3),
        0
    )

    _, thresh = cv2.threshold(
        gray,
        150,
        255,
        cv2.THRESH_BINARY
    )

    processed_path = "cap_processed.png"

    cv2.imwrite(
        processed_path,
        thresh
    )

    img = Image.open(
        processed_path
    )

    captcha_text = pytesseract.image_to_string(
        img,
        config=(
            "--oem 3 "
            "--psm 7 "
            "-c tessedit_char_whitelist="
            "abcdefghijklmnopqrstuvwxyz"
            "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
            "0123456789"
        )
    )

    return clean_captcha(
        captcha_text
    )


# ============================================================
# FETCH SINGLE RESULT
# ============================================================

def fetch_result(usn):
    exact_codes = get_semester_subject_codes(semc)
    target_prefix = get_target_subject_prefix(semc)

    attempt = 0
    while True:
        attempt += 1

        try:
            print(f"\n🔎 {usn} - Attempt {attempt} (unlimited retries)")

            # ------------------------------------------------
            # OPEN RESULT PAGE
            # ------------------------------------------------

            try:
                driver.get(url)
            except TimeoutException:
                print(
                    f"⚠️ Page loading timed out for {usn}."
                )
                try:
                    driver.execute_script(
                        "window.stop();"
                    )
                except Exception:
                    pass

            WebDriverWait(
                driver,
                ELEMENT_TIMEOUT
            ).until(
                EC.presence_of_element_located(
                    (By.NAME, "lns")
                )
            )

            # ------------------------------------------------
            # ENTER USN
            # ------------------------------------------------

            usn_input = driver.find_element(
                By.NAME,
                "lns"
            )

            usn_input.clear()
            usn_input.send_keys(usn)

            # ------------------------------------------------
            # CAPTCHA
            # ------------------------------------------------

            captcha_img = WebDriverWait(
                driver,
                ELEMENT_TIMEOUT
            ).until(
                EC.presence_of_element_located(
                    (
                        By.XPATH,
                        "//img[@alt='CAPTCHA code']"
                    )
                )
            )

            captcha_img.screenshot(
                "cap.png"
            )

            captcha_text = solve_captcha(
                "cap.png"
            )

            if len(captcha_text) < 4:
                print(
                    "⚠️ OCR could not read the CAPTCHA."
                )

                time.sleep(RETRY_DELAY)
                continue

            print(
                f"🔐 OCR CAPTCHA for {usn}: "
                f"{captcha_text}"
            )

            # ------------------------------------------------
            # ENTER CAPTCHA
            # ------------------------------------------------

            cap_input = driver.find_element(
                By.NAME,
                "captchacode"
            )

            cap_input.clear()
            cap_input.send_keys(
                captcha_text
            )

            # ------------------------------------------------
            # SUBMIT
            # ------------------------------------------------

            driver.find_element(
                By.ID,
                "submit"
            ).click()

            time.sleep(2)

            # ------------------------------------------------
            # ALERTS
            # ------------------------------------------------

            try:
                alert = driver.switch_to.alert
                alert_text = alert.text

                alert.accept()

                if "Invalid captcha" in alert_text:
                    print(
                        "❌ Invalid CAPTCHA."
                    )

                    time.sleep(RETRY_DELAY)
                    continue

                if "University Seat Number" in alert_text:
                    print(
                        f"⚠️ No result currently available for {usn}; retrying."
                    )
                    time.sleep(RETRY_DELAY)
                    continue

            except Exception:
                pass

            # ------------------------------------------------
            # PARSE HTML
            # ------------------------------------------------

            soup = BeautifulSoup(
                driver.page_source,
                "html.parser"
            )

            with open(
                "page.html",
                "w",
                encoding="utf-8"
            ) as file:
                file.write(
                    driver.page_source
                )

            student = {}

            # ------------------------------------------------
            # STUDENT DETAILS
            # ------------------------------------------------

            info_rows = soup.select(
                ".divTableRow"
            )

            for row in info_rows:

                cells = row.select(
                    ".divTableCell"
                )

                if len(cells) != 2:
                    continue

                key = cells[0].get_text(
                    " ",
                    strip=True
                )

                value = cells[1].get_text(
                    " ",
                    strip=True
                )

                if "University Seat Number" in key:
                    student["USN"] = value

                elif "Student Name" in key:
                    student["Name"] = value

            if "USN" not in student:
                print(
                    f"⚠️ Could not read student details "
                    f"for {usn}."
                )

                time.sleep(RETRY_DELAY)
                continue

            # ------------------------------------------------
            # SEMESTER
            # ------------------------------------------------

            semester_text = ""

            sem_tag = soup.find(
                "b",
                string=lambda x:
                    x and "Semester" in x
            )

            if sem_tag:
                semester_text = sem_tag.get_text(
                    strip=True
                )

            student["Semester"] = semester_text

            # ------------------------------------------------
            # SUBJECT DETAILS
            # ------------------------------------------------

            subjects = []

            rows = soup.find_all(
                "div",
                class_="divTableRow"
            )

            if exact_codes:
                print(
                    f"🎯 {usn}: extracting exact "
                    f"Semester {semc} subjects"
                )
            else:
                print(
                    f"🎯 {usn}: extracting "
                    f"{target_prefix} subjects"
                )

            for row in rows:

                cells = row.find_all(
                    "div",
                    class_="divTableCell"
                )

                if len(cells) != 7:
                    continue

                code = cells[0].get_text(
                    strip=True
                ).upper()

                if code == "SUBJECT CODE":
                    continue

                if code.startswith("P ->"):
                    continue

                # EXACT SEMESTER FILTER
                if not is_subject_for_semester(
                    code,
                    semc
                ):
                    continue

                subject = {
                    "code": code,
                    "name": cells[1].get_text(
                        strip=True
                    ),
                    "internal": cells[2].get_text(
                        strip=True
                    ),
                    "external": cells[3].get_text(
                        strip=True
                    ),
                    "total": cells[4].get_text(
                        strip=True
                    ),
                    "result": cells[5].get_text(
                        strip=True
                    ),
                    "date": cells[6].get_text(
                        strip=True
                    )
                }

                subjects.append(
                    subject
                )

            # ------------------------------------------------
            # ORDER SUBJECTS
            # ------------------------------------------------

            if exact_codes:
                subject_map = {
                    subject["code"]: subject
                    for subject in subjects
                }

                ordered_subjects = []

                for code in exact_codes:
                    if code in subject_map:
                        ordered_subjects.append(
                            subject_map[code]
                        )

                subjects = ordered_subjects

            student["subjects"] = subjects

            print(
                f"📚 {usn}: "
                f"{len(subjects)} subjects selected"
            )

            if not subjects:
                print(
                    f"⚠️ No Semester {semc} subjects "
                    f"found for {usn}."
                )

                time.sleep(RETRY_DELAY)
                continue

            # ------------------------------------------------
            # SUCCESS
            # ------------------------------------------------

            print(
                f"✅ Result fetched successfully "
                f"for {usn}"
            )

            return student

        except (
            TimeoutException,
            WebDriverException,
            ConnectionError,
            OSError
        ) as e:

            print(
                f"⚠️ Browser/network error for "
                f"{usn}: {type(e).__name__}"
            )

            print(f"🔄 Retrying {usn} (attempt {attempt + 1})...")
            time.sleep(RETRY_DELAY)
            continue

        except Exception as e:

            print(
                f"⚠️ Error fetching {usn}: {e}"
            )

            print(f"🔄 Retrying {usn} (attempt {attempt + 1})...")
            time.sleep(RETRY_DELAY)
            continue

    return None


# ============================================================
# SAFE NUMBER
# ============================================================

def safe_number(value):
    if value is None:
        return 0

    text = str(value).strip()

    if not text:
        return 0

    try:
        return float(text)
    except Exception:
        return 0


# ============================================================
# CLASS
# ============================================================

def get_class(percentage):

    if percentage >= 70:
        return "FCD"

    if percentage >= 60:
        return "FC"

    if percentage >= 50:
        return "SC"

    if percentage >= 40:
        return "PASS"

    return "FAIL"


# ============================================================
# SAVE RESULTS TO EXCEL
# ============================================================

def save_results_to_excel(results):

    if not results:
        print("❌ No Results Found")
        return

    wb = Workbook()
    ws = wb.active
    ws.title = "VTU Results"

    # --------------------------------------------------------
    # COLORS
    # --------------------------------------------------------

    dark_blue = "1F2937"
    subject_blue = "D9EAF7"
    white = "FFFFFF"
    black = "000000"
    red = "FF0000"
    light_green = "D9EAD3"

    dark_fill = PatternFill(
        fill_type="solid",
        fgColor=dark_blue
    )

    subject_fill = PatternFill(
        fill_type="solid",
        fgColor=subject_blue
    )

    red_fill = PatternFill(
        fill_type="solid",
        fgColor=red
    )

    green_fill = PatternFill(
        fill_type="solid",
        fgColor=light_green
    )

    header_font = Font(
        bold=True,
        color=white,
        size=10
    )

    subject_font = Font(
        bold=True,
        color=black,
        size=9
    )

    normal_font = Font(
        size=9
    )

    red_font = Font(
        bold=True,
        color=white
    )

    green_font = Font(
        bold=True,
        color="006100"
    )

    thin_side = Side(
        style="thin",
        color="000000"
    )

    thin_border = Border(
        left=thin_side,
        right=thin_side,
        top=thin_side,
        bottom=thin_side
    )

    center = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True
    )

    left = Alignment(
        horizontal="left",
        vertical="center",
        wrap_text=True
    )

    # --------------------------------------------------------
    # SUBJECT LIST
    # --------------------------------------------------------

    exact_codes = get_semester_subject_codes(
        semc
    )

    all_subjects = {}

    for student in results:
        for subject in student.get(
            "subjects",
            []
        ):

            code = str(
                subject.get(
                    "code",
                    ""
                )
            ).strip().upper()

            if not is_subject_for_semester(
                code,
                semc
            ):
                continue

            if code not in all_subjects:
                all_subjects[code] = subject

    if exact_codes:
        subjects = [
            all_subjects[code]
            for code in exact_codes
            if code in all_subjects
        ]
    else:
        subjects = list(
            all_subjects.values()
        )

    subject_count = len(subjects)

    print(
        f"📊 Excel will contain "
        f"{subject_count} selected-semester subjects."
    )

    if exact_codes:
        print(
            "🎯 Required codes: "
            + ", ".join(exact_codes)
        )

    if subject_count == 0:
        print(
            "❌ No selected-semester subjects "
            "available for Excel."
        )
        return

    # --------------------------------------------------------
    # BASIC HEADERS
    # --------------------------------------------------------

    ws["A1"] = "Sl.\nNo."
    ws["B1"] = "USN"
    ws["C1"] = "Name"

    ws.merge_cells("A1:A2")
    ws.merge_cells("B1:B2")
    ws.merge_cells("C1:C2")

    # --------------------------------------------------------
    # SUBJECT HEADERS
    # --------------------------------------------------------

    current_col = 4

    for subject in subjects:

        start_col = current_col
        end_col = current_col + 3

        start_letter = get_column_letter(
            start_col
        )

        end_letter = get_column_letter(
            end_col
        )

        subject_code = str(
            subject.get(
                "code",
                ""
            )
        ).strip()

        subject_name = str(
            subject.get(
                "name",
                ""
            )
        ).strip()

        header_text = (
            f"{subject_code}\n"
            f"{subject_name}"
            if subject_name
            else subject_code
        )

        ws.merge_cells(
            f"{start_letter}1:"
            f"{end_letter}1"
        )

        header_cell = ws.cell(
            row=1,
            column=start_col
        )

        header_cell.value = header_text
        header_cell.fill = subject_fill
        header_cell.font = subject_font
        header_cell.alignment = center
        header_cell.border = thin_border

        sub_headers = [
            "INT",
            "EXT",
            "T",
            "R"
        ]

        for offset, title in enumerate(
            sub_headers
        ):

            cell = ws.cell(
                row=2,
                column=current_col + offset
            )

            cell.value = title
            cell.fill = dark_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = thin_border

        current_col += 4

    # --------------------------------------------------------
    # FINAL COLUMNS
    # --------------------------------------------------------

    total_col = current_col
    percentage_col = current_col + 1
    class_col = current_col + 2
    remark_col = current_col + 3

    final_headers = {
        total_col: "Total",
        percentage_col: "%",
        class_col: "Class",
        remark_col: "Remark"
    }

    for col, value in final_headers.items():

        cell = ws.cell(
            row=1,
            column=col
        )

        cell.value = value
        cell.fill = dark_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border

        letter = get_column_letter(
            col
        )

        ws.merge_cells(
            f"{letter}1:{letter}2"
        )

    # --------------------------------------------------------
    # BASIC HEADER STYLE
    # --------------------------------------------------------

    for cell in [
        ws["A1"],
        ws["B1"],
        ws["C1"]
    ]:

        cell.fill = dark_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = thin_border

    # --------------------------------------------------------
    # STUDENTS
    # --------------------------------------------------------

    for student_index, student in enumerate(
        results,
        start=1
    ):

        excel_row = student_index + 2

        ws.cell(
            row=excel_row,
            column=1
        ).value = student_index

        ws.cell(
            row=excel_row,
            column=2
        ).value = student.get(
            "USN",
            ""
        )

        ws.cell(
            row=excel_row,
            column=3
        ).value = student.get(
            "Name",
            ""
        )

        total_marks = 0
        failed_student = False

        current_col = 4

        student_subjects = {
            str(
                subject.get(
                    "code",
                    ""
                )
            ).upper(): subject
            for subject in student.get(
                "subjects",
                []
            )
            if is_subject_for_semester(
                subject.get("code", ""),
                semc
            )
        }

        for subject_header in subjects:

            code_key = str(
                subject_header.get(
                    "code",
                    ""
                )
            ).upper()

            subject = student_subjects.get(
                code_key,
                {}
            )

            internal = subject.get(
                "internal",
                ""
            )

            external = subject.get(
                "external",
                ""
            )

            total = subject.get(
                "total",
                ""
            )

            result = str(
                subject.get(
                    "result",
                    ""
                )
            ).strip().upper()

            total_marks += safe_number(
                total
            )

            if result == "F":
                failed_student = True

            values = [
                internal,
                external,
                total,
                result
            ]

            for offset, value in enumerate(
                values
            ):

                cell = ws.cell(
                    row=excel_row,
                    column=current_col + offset
                )

                cell.value = value
                cell.alignment = center

                if result == "F":
                    cell.fill = red_fill
                    cell.font = red_font

            current_col += 4

        # ----------------------------------------------------
        # PERCENTAGE
        # ----------------------------------------------------

        maximum_marks = subject_count * 100

        percentage = (
            (total_marks / maximum_marks) * 100
            if maximum_marks > 0
            else 0
        )

        percentage = round(
            percentage,
            2
        )

        # ----------------------------------------------------
        # CLASS
        # ----------------------------------------------------

        student_class = get_class(
            percentage
        )

        # ----------------------------------------------------
        # REMARK
        # ----------------------------------------------------

        remark = (
            "FAIL"
            if failed_student
            else "PASS"
        )

        # ----------------------------------------------------
        # FINAL COLUMNS
        # ----------------------------------------------------

        ws.cell(
            row=excel_row,
            column=total_col
        ).value = int(
            round(total_marks)
        )

        ws.cell(
            row=excel_row,
            column=percentage_col
        ).value = percentage

        ws.cell(
            row=excel_row,
            column=class_col
        ).value = student_class

        ws.cell(
            row=excel_row,
            column=remark_col
        ).value = remark

        # Style final cells

        for col in [
            total_col,
            percentage_col,
            class_col,
            remark_col
        ]:

            cell = ws.cell(
                row=excel_row,
                column=col
            )

            cell.alignment = center

        if failed_student:

            for col in [
                total_col,
                class_col,
                remark_col
            ]:

                cell = ws.cell(
                    row=excel_row,
                    column=col
                )

                cell.fill = red_fill
                cell.font = red_font

        else:

            remark_cell = ws.cell(
                row=excel_row,
                column=remark_col
            )

            remark_cell.fill = green_fill
            remark_cell.font = green_font

    # --------------------------------------------------------
    # STYLE DATA
    # --------------------------------------------------------

    for row in ws.iter_rows(
        min_row=3,
        max_row=ws.max_row,
        min_col=1,
        max_col=remark_col
    ):

        for cell in row:

            cell.border = thin_border

            if cell.column == 3:
                cell.alignment = left
            else:
                cell.alignment = center

            if cell.font == Font():
                cell.font = normal_font

    # --------------------------------------------------------
    # WIDTHS
    # --------------------------------------------------------

    ws.column_dimensions["A"].width = 7
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 30

    current_col = 4

    for _ in subjects:

        ws.column_dimensions[
            get_column_letter(current_col)
        ].width = 8

        ws.column_dimensions[
            get_column_letter(current_col + 1)
        ].width = 8

        ws.column_dimensions[
            get_column_letter(current_col + 2)
        ].width = 8

        ws.column_dimensions[
            get_column_letter(current_col + 3)
        ].width = 7

        current_col += 4

    ws.column_dimensions[
        get_column_letter(total_col)
    ].width = 10

    ws.column_dimensions[
        get_column_letter(percentage_col)
    ].width = 10

    ws.column_dimensions[
        get_column_letter(class_col)
    ].width = 12

    ws.column_dimensions[
        get_column_letter(remark_col)
    ].width = 12

    # --------------------------------------------------------
    # ROWS
    # --------------------------------------------------------

    ws.row_dimensions[1].height = 55
    ws.row_dimensions[2].height = 25

    for row in range(
        3,
        ws.max_row + 1
    ):
        ws.row_dimensions[row].height = 22

    # --------------------------------------------------------
    # FREEZE + FILTER
    # --------------------------------------------------------

    ws.freeze_panes = "D3"

    last_column_letter = get_column_letter(
        remark_col
    )

    ws.auto_filter.ref = (
        f"A2:{last_column_letter}"
        f"{ws.max_row}"
    )

    # --------------------------------------------------------
    # PAGE SETUP
    # --------------------------------------------------------

    ws.sheet_view.showGridLines = True

    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # --------------------------------------------------------
    # SAVE
    # --------------------------------------------------------

    os.makedirs(
        "ExcelFiles",
        exist_ok=True
    )

    output_file = os.path.join(
        "ExcelFiles",
        "results.xlsx"
    )

    wb.save(output_file)

    print("\n==========================================")
    print("✅ Excel Saved Successfully")
    print(f"📁 File: {output_file}")
    print(f"🎓 Semester: {semc}")
    print(f"📚 Subjects Included: {subject_count}")
    print("==========================================\n")


# ============================================================
# RANGE SCRAPER
# ============================================================

def scrape_range():

    results = []

    for number in range(
        low,
        high
    ):

        usn = (
            f"{college}"
            f"{year}"
            f"{branch}"
            f"{str(number).zfill(3)}"
        )

        print("\n================================")
        print(
            f"Fetching result for {usn}"
        )
        print("================================")

        result = fetch_result(
            usn
        )

        if result:
            results.append(
                result
            )

        else:
            print(
                f"⏭️ Skipping {usn}."
            )

    print(
        f"\nTotal Results: {len(results)}"
    )

    if results:
        save_results_to_excel(
            results
        )
    else:
        print(
            "❌ No successful results were found."
        )

        # Remove stale Excel if one exists
        output_file = os.path.join(
            "ExcelFiles",
            "results.xlsx"
        )

        if os.path.exists(output_file):
            try:
                os.remove(output_file)
            except Exception:
                pass


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":

    (
        college,
        year,
        branch,
        low,
        high,
        semc,
        url
    ) = get_inputs()

    print("\n==========================================")
    print("       VTU SEMESTER RESULT SCRAPER")
    print("==========================================")
    print(f"College : {college}")
    print(f"Year    : {year}")
    print(f"Branch  : {branch}")
    print(f"USN     : {low} - {high - 1}")
    print(f"Semester: {semc}")

    exact_codes = get_semester_subject_codes(
        semc
    )

    if exact_codes:
        print(
            "Subjects: "
            + ", ".join(exact_codes)
        )
    else:
        print(
            "Filter  : "
            + get_target_subject_prefix(semc)
        )

    print("==========================================\n")

    os.makedirs(
        "captcha_dataset",
        exist_ok=True
    )

    try:

        driver = create_driver()

        scrape_range()

    except KeyboardInterrupt:

        print(
            "\n⛔ Scraper stopped by user."
        )

    except Exception as e:

        print(
            f"\n❌ Scraper stopped due to error: {e}"
        )

    finally:

        if driver is not None:

            try:
                driver.quit()
                print(
                    "🧹 Chrome driver closed."
                )
            except Exception:
                pass

        print(
            "🎉 SCRAPING PROCESS FINISHED"
        )
