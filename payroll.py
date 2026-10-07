"""
Automated Multi-Page Payslip Emailer
====================================

Assumption:
    Each page in an input PDF is one employee's payslip.

Workflow:
1. Read a multi-page PDF from the input folder.
2. Split it into individual one-page PDFs.
3. Extract the Employee ID from each page.
4. Use OCR when a page has no searchable text.
5. Match the Employee ID against the employee Excel file.
6. Password-protect that employee's one-page payslip using the IC number.
7. Email only that page to the matching employee.
8. Archive the protected one-page PDF.
9. Delete the original multi-page source only when every page was sent.
10. Keep the source PDF in input when one or more pages need attention.

Required Python packages:
    pip install pandas openpyxl pypdf pymupdf pillow pytesseract

Tesseract OCR must also be installed on Windows.
Common installation path:
    C:\\Program Files\\Tesseract-OCR\\tesseract.exe
"""

import hashlib
import io
import os
import re
import shutil
import smtplib
import traceback

from datetime import datetime
from email.message import EmailMessage
from getpass import getpass
from pathlib import Path

import pymupdf
import pandas as pd
import pytesseract

from PIL import Image, ImageOps
from pypdf import PdfReader, PdfWriter

# ============================================================
# CONFIGURATION
# ============================================================

BASE_PATH = Path(r"C:\Repo\payslip")

PDF_FOLDER = BASE_PATH / "input"
ARCHIVE_FOLDER = BASE_PATH / "archive"
MANUAL_REVIEW_FOLDER = BASE_PATH / "manual_review"
TEMP_FOLDER = BASE_PATH / "temp"
SPLIT_FOLDER = TEMP_FOLDER / "split_pages"
PROTECTED_FOLDER = TEMP_FOLDER / "protected_pages"

DATA_FOLDER = BASE_PATH / "data"
EXTRACTED_TEXT_FOLDER = DATA_FOLDER / "extracted_text"

EXCEL_FILE = DATA_FOLDER / "payroll_employee_data.xlsx"
LOG_FILE = DATA_FOLDER / "email_log.xlsx"
ERROR_LOG = DATA_FOLDER / "payroll_error.log"

SENDER_EMAIL = "payroll@bses.com.my"
SMTP_SERVER = "smtp.office365.com"
SMTP_PORT = 587

# Save extracted PDF/OCR text for troubleshooting.
SAVE_EXTRACTED_TEXT = True

# Delete the original multi-page PDF only after all pages have
# a SUCCESS record in the log.
DELETE_SOURCE_AFTER_ALL_PAGES_SENT = True

# Minimum number of alphanumeric characters required before a PDF's
# normal text layer is treated as usable.
MINIMUM_TEXT_CHARACTERS = 30

# Update this path if Tesseract is installed somewhere else.
TESSERACT_CMD = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")

if TESSERACT_CMD.exists():
    pytesseract.pytesseract.tesseract_cmd = str(TESSERACT_CMD)


# ============================================================
# GENERAL HELPERS
# ============================================================


def extract_payslip_period(text):
    """
    Extract payslip period from text such as:
    01.03.2026 - 31.03.2026

    Returns:
        March 2026
    """

    match = re.search(r"\b\d{2}\.(\d{2})\.(20\d{2})\s*-\s*\d{2}\.\d{2}\.\d{4}\b", text)

    if not match:
        raise ValueError("Unable to detect payslip period from payslip text.")

    month_number = int(match.group(1))
    year = match.group(2)

    month_name = datetime.strptime(str(month_number), "%m").strftime("%B")

    return f"{month_name} {year}"


def ensure_folders():
    """Create all required folders."""

    for folder in (
        PDF_FOLDER,
        ARCHIVE_FOLDER,
        MANUAL_REVIEW_FOLDER,
        TEMP_FOLDER,
        SPLIT_FOLDER,
        PROTECTED_FOLDER,
        DATA_FOLDER,
        EXTRACTED_TEXT_FOLDER,
    ):
        folder.mkdir(parents=True, exist_ok=True)


def normalize_header(value):
    """
    Convert an Excel heading into a standard comparable format.

    Examples:
        Employee ID -> employeeid
        employee_ID -> employeeid
        IC Number   -> icnumber
    """

    if pd.isna(value):
        return ""

    return re.sub(
        r"[^a-z0-9]",
        "",
        str(value).strip().lower(),
    )


def clean_excel_value(value):
    """Convert an Excel value into a clean string."""

    if pd.isna(value):
        return ""

    value = str(value).strip()

    # Remove .0 added when Excel stores an identifier as a number.
    value = re.sub(r"\.0$", "", value)

    return value


def normalize_employee_id(value):
    """Keep only digits in an Employee ID."""

    return re.sub(
        r"\D",
        "",
        clean_excel_value(value),
    )


def valid_email(value):
    """Perform a basic email-address validation."""

    return bool(
        re.fullmatch(
            r"[^@\s]+@[^@\s]+\.[^@\s]+",
            clean_excel_value(value),
        )
    )


def calculate_file_hash(file_path):
    """Return the SHA-256 hash of a file."""

    digest = hashlib.sha256()

    with open(file_path, "rb") as file_handle:
        for block in iter(
            lambda: file_handle.read(1024 * 1024),
            b"",
        ):
            digest.update(block)

    return digest.hexdigest()


def unique_destination(folder, filename):
    """Return a path that will not overwrite an existing file."""

    folder = Path(folder)
    filename = Path(filename)

    candidate = folder / filename.name

    if not candidate.exists():
        return candidate

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")

    return folder / (f"{filename.stem}_{timestamp}{filename.suffix}")


def safe_filename_part(value):
    """Remove characters that are unsafe in Windows filenames."""

    cleaned = re.sub(
        r'[<>:"/\\|?*]+',
        "_",
        str(value),
    )

    cleaned = re.sub(r"\s+", " ", cleaned).strip()

    return cleaned or "payslip"


def write_error_log(message):
    """Append an error and traceback details to the error log."""

    DATA_FOLDER.mkdir(parents=True, exist_ok=True)

    with open(
        ERROR_LOG,
        "a",
        encoding="utf-8",
    ) as file_handle:

        file_handle.write("\n\n")
        file_handle.write("=" * 70 + "\n")
        file_handle.write(
            datetime.now().isoformat(
                sep=" ",
                timespec="seconds",
            )
        )
        file_handle.write("\n")
        file_handle.write(str(message))
        file_handle.write("\n")
        traceback.print_exc(file=file_handle)


# ============================================================
# EMPLOYEE EXCEL DATA
# ============================================================


def load_employee_data():
    """
    Load and validate employee data.

    Required logical columns:
        Employee ID
        Name
        Email
        IC Number

    The header may be within the first 30 Excel rows.
    """

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Excel file not found:\n{EXCEL_FILE}")

    column_aliases = {
        # Employee ID
        "employeeid": "Employee ID",
        "empid": "Employee ID",
        "staffid": "Employee ID",
        "employeenumber": "Employee ID",
        "staffnumber": "Employee ID",
        # Employee name
        "employeename": "Name",
        "staffname": "Name",
        "name": "Name",
        # Email
        "email": "Email",
        "emailaddress": "Email",
        "employeeemail": "Email",
        "staffemail": "Email",
        # IC number
        "ic": "IC Number",
        "icnumber": "IC Number",
        "nric": "IC Number",
        "nricnumber": "IC Number",
        "identitycard": "IC Number",
        "identitycardnumber": "IC Number",
    }

    normalized_aliases = {
        normalize_header(alias): standard_name
        for alias, standard_name in column_aliases.items()
    }

    required_columns = {
        "Employee ID",
        "Name",
        "Email",
        "IC Number",
    }

    preview = pd.read_excel(
        EXCEL_FILE,
        header=None,
        nrows=30,
        dtype=str,
    )

    header_row = None

    for row_number, row in preview.iterrows():
        detected_columns = set()

        for cell in row:
            normalized_cell = normalize_header(cell)

            if normalized_cell in normalized_aliases:
                detected_columns.add(normalized_aliases[normalized_cell])

        if required_columns.issubset(detected_columns):
            header_row = row_number
            break

    if header_row is None:
        raise ValueError(
            "Unable to find the Excel header row.\n\n"
            "The Excel file must contain headings for:\n"
            "- Employee ID\n"
            "- Name\n"
            "- Email\n"
            "- IC Number"
        )

    print(f"Excel header detected on row {header_row + 1}")

    employees = pd.read_excel(
        EXCEL_FILE,
        header=header_row,
        dtype=str,
    )

    employees.dropna(
        axis=1,
        how="all",
        inplace=True,
    )

    employees.dropna(
        axis=0,
        how="all",
        inplace=True,
    )

    rename_columns = {}

    for column in employees.columns:
        normalized_column = normalize_header(column)

        if normalized_column in normalized_aliases:
            rename_columns[column] = normalized_aliases[normalized_column]

    employees.rename(
        columns=rename_columns,
        inplace=True,
    )

    missing_columns = [
        column for column in required_columns if column not in employees.columns
    ]

    if missing_columns:
        raise ValueError(
            "Missing required Excel column(s): " + ", ".join(missing_columns)
        )

    employees = employees[
        [
            "Employee ID",
            "Name",
            "Email",
            "IC Number",
        ]
    ].copy()

    for column in employees.columns:
        employees[column] = employees[column].apply(clean_excel_value)

    employees["Employee ID"] = employees["Employee ID"].apply(normalize_employee_id)

    employees = employees[employees["Employee ID"] != ""].copy()

    duplicate_ids = employees[employees["Employee ID"].duplicated(keep=False)]

    if not duplicate_ids.empty:
        duplicate_list = sorted(duplicate_ids["Employee ID"].unique().tolist())

        raise ValueError(
            "Duplicate Employee ID values were found in Excel:\n"
            + "\n".join(duplicate_list)
        )

    return employees


# ============================================================
# PDF SPLITTING
# ============================================================


def create_single_page_pdf(
    source_reader,
    page_index,
    output_pdf,
):
    """
    Write one page from a source PdfReader into a new PDF.

    page_index is zero-based.
    """

    writer = PdfWriter()
    writer.add_page(source_reader.pages[page_index])

    with open(output_pdf, "wb") as file_handle:
        writer.write(file_handle)


# ============================================================
# PDF TEXT EXTRACTION AND OCR
# ============================================================


def usable_text_score(text):
    """Count usable alphanumeric characters."""

    return len(
        re.findall(
            r"[A-Za-z0-9]",
            text or "",
        )
    )


def extract_normal_pdf_text(pdf_path):
    """Extract text from a PDF's searchable text layer."""

    reader = PdfReader(str(pdf_path))

    return "\n".join(page.extract_text() or "" for page in reader.pages)


def score_ocr_text(text):
    """Score OCR output and reward payslip headings."""

    score = usable_text_score(text)

    heading_patterns = (
        r"\bEMP\s*\.?\s*#",
        r"\bEMPLOYEE\s*ID\b",
        r"\bNAME\s*:",
        r"\bTOTAL\s+EARNINGS\b",
        r"\bNETT?\s+WAGE\b",
    )

    for pattern in heading_patterns:
        if re.search(
            pattern,
            text or "",
            re.IGNORECASE,
        ):
            score += 1000

    return score


def verify_tesseract():
    """Check that Tesseract OCR is available."""

    try:
        pytesseract.get_tesseract_version()
    except Exception as error:
        raise RuntimeError(
            "Tesseract OCR is not installed or cannot be found.\n"
            "Install Tesseract and update TESSERACT_CMD in this script."
        ) from error


def ocr_pdf_text(pdf_path):
    """
    OCR a one-page or multi-page image PDF.

    Every page is rendered at 300 DPI and tested at:
        0, 90, 180, and 270 degrees.
    """

    verify_tesseract()

    document = pymupdf.open(str(pdf_path))
    page_results = []

    try:
        for page_number, page in enumerate(
            document,
            start=1,
        ):
            pixmap = page.get_pixmap(
                dpi=300,
                alpha=False,
            )

            image = Image.open(io.BytesIO(pixmap.tobytes("png")))

            image = ImageOps.grayscale(image)
            image = ImageOps.autocontrast(image)

            best_text = ""
            best_score = -1
            best_angle = 0

            for angle in (0, 90, 180, 270):
                rotated_image = image.rotate(
                    angle,
                    expand=True,
                )

                candidate_text = pytesseract.image_to_string(
                    rotated_image,
                    config="--oem 3 --psm 6",
                    lang="eng",
                )

                candidate_score = score_ocr_text(candidate_text)

                if candidate_score > best_score:
                    best_score = candidate_score
                    best_text = candidate_text
                    best_angle = angle

            print(f"OCR page {page_number}: " f"selected rotation {best_angle} degrees")

            page_results.append(best_text)

    finally:
        document.close()

    return "\n".join(page_results)


def extract_pdf_text(pdf_path):
    """
    Use normal PDF text extraction first.
    Use OCR only when the normal text layer is unusable.

    Returns:
        text, extraction_method
    """

    normal_text = extract_normal_pdf_text(pdf_path)

    if usable_text_score(normal_text) >= MINIMUM_TEXT_CHARACTERS:
        return normal_text, "PDF_TEXT"

    print("No usable searchable text found. " "Starting OCR fallback...")

    ocr_text = ocr_pdf_text(pdf_path)

    return ocr_text, "OCR"


def save_extracted_text(
    source_name,
    page_number,
    text,
    method,
):
    """Save one page's extracted text for troubleshooting."""

    if not SAVE_EXTRACTED_TEXT:
        return

    safe_source = safe_filename_part(Path(source_name).stem)

    output_name = f"{safe_source}" f"_page_{page_number:03d}" f"_{method}.txt"

    output_path = EXTRACTED_TEXT_FOLDER / output_name

    with open(
        output_path,
        "w",
        encoding="utf-8",
    ) as file_handle:
        file_handle.write(text or "")


# ============================================================
# EMPLOYEE-ID DETECTION
# ============================================================


def extract_employee_id(text):
    """
    Extract an Employee ID from a payslip page.

    At least five digits are required.
    """

    cleaned_text = (text or "").replace("\u00a0", " ")

    cleaned_text = re.sub(
        r"\s+",
        " ",
        cleaned_text,
    )

    patterns = (
        r"\bEmployee\s*(?:ID|No|Number)" r"\s*[:#\-]?\s*(\d{5,})\b",
        r"\bEMP\s*\.?\s*#" r"\s*[:\-]?\s*(\d{5,})\b",
        r"\bEMP\s*(?:ID|No|Number)" r"\s*[:#\-]?\s*(\d{5,})\b",
    )

    for pattern in patterns:
        result = re.search(
            pattern,
            cleaned_text,
            re.IGNORECASE,
        )

        if result:
            return normalize_employee_id(result.group(1))

    return None


def find_employee(emp_id, employees):
    """Return exactly one matching employee row, or None."""

    if not emp_id:
        return None

    matches = employees[employees["Employee ID"] == emp_id]

    if len(matches) != 1:
        return None

    return matches.iloc[0]


# ============================================================
# PDF PASSWORD PROTECTION
# ============================================================


def protect_pdf(
    input_pdf,
    output_pdf,
    password,
):
    """Create a password-protected copy of a PDF."""

    password = clean_excel_value(password)

    if not password:
        raise ValueError(
            "The employee IC number is empty; " "the PDF cannot be password-protected."
        )

    reader = PdfReader(str(input_pdf))
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(password)

    with open(
        output_pdf,
        "wb",
    ) as file_handle:
        writer.write(file_handle)


# ============================================================
# LOGGING
# ============================================================

LOG_COLUMNS = [
    "Date",
    "Employee ID",
    "Email",
    "Status",
    "Source File",
    "Page Number",
    "Page Key",
    "File Hash",
    "Extraction Method",
    "Archived File",
    "Details",
]


def load_log():
    """Load the processing log and add any missing columns."""

    if LOG_FILE.exists():
        log = pd.read_excel(
            LOG_FILE,
            dtype=str,
        )
    else:
        log = pd.DataFrame(columns=LOG_COLUMNS)

    for column in LOG_COLUMNS:
        if column not in log.columns:
            log[column] = ""

    log = log[LOG_COLUMNS].copy()

    for column in log.columns:
        log[column] = log[column].fillna("").astype(str).str.strip()

    log["Employee ID"] = log["Employee ID"].apply(normalize_employee_id)

    return log


def save_log(log):
    """Save the processing log."""

    log.to_excel(
        LOG_FILE,
        index=False,
    )


def append_log(
    log,
    emp_id,
    email,
    status,
    source_file,
    page_number,
    page_key,
    file_hash,
    extraction_method,
    archived_file,
    details,
):
    """Append one page-processing result to the log."""

    log.loc[len(log)] = {
        "Date": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "Employee ID": emp_id or "",
        "Email": email or "",
        "Status": status,
        "Source File": source_file,
        "Page Number": str(page_number),
        "Page Key": page_key,
        "File Hash": file_hash,
        "Extraction Method": extraction_method or "",
        "Archived File": archived_file or "",
        "Details": details or "",
    }

    save_log(log)


def page_already_sent(page_key, log):
    """Check whether this page was already emailed successfully."""

    successful = log[
        (log["Page Key"] == page_key) & (log["Status"].str.upper() == "SUCCESS")
    ]

    return not successful.empty


def every_page_sent(
    source_hash,
    total_pages,
    log,
):
    """Return True only when every source page has a SUCCESS record."""

    for page_number in range(
        1,
        total_pages + 1,
    ):
        page_key = f"{source_hash}:page:{page_number}"

        if not page_already_sent(
            page_key,
            log,
        ):
            return False

    return True


# ============================================================
# EMAIL
# ============================================================


def test_smtp_login(outlook_password):
    """
    Test Microsoft 365 authentication before processing PDFs.

    This prevents OCR work and repeated page failures when the mailbox
    password or SMTP AUTH configuration is incorrect.
    """

    print("Testing Microsoft 365 SMTP login...")

    with smtplib.SMTP(
        SMTP_SERVER,
        SMTP_PORT,
        timeout=60,
    ) as server:

        server.ehlo()
        server.starttls()
        server.ehlo()

        server.login(
            SENDER_EMAIL,
            outlook_password,
        )

    print("SMTP login successful.")


def send_email(
    email,
    emp_id,
    name,
    pdf_path,
    outlook_password,
    payslip_period,
):
    """Send one protected one-page payslip."""

    if not valid_email(email):
        raise ValueError("Invalid or empty employee " f"email address: {email!r}")

    message = EmailMessage()

    message["From"] = SENDER_EMAIL
    message["To"] = email
    message["Subject"] = f"e-Payslip for {payslip_period}"

    # Plain-text fallback for email clients that do not display HTML.
    message.set_content(f"""PRIVATE & CONFIDENTIAL

Dear {name.upper()},

Attached is your e-Payslip for {payslip_period}.
Password: Your IC number without dashes.
Example: If your IC number is 980828-12-6128, the password is 980828126128.

This is an auto-generated email. Please do not reply.

Kind regards,

Payroll Unit
""")

    # HTML version allows the confidentiality notice to appear bold and red.
    message.add_alternative(
        f"""<!DOCTYPE html>
<html>
<body style="font-family: Arial, sans-serif; font-size: 14px; color: #000000;">
    <p style="font-weight: bold; color: red;">PRIVATE &amp; CONFIDENTIAL</p>

    <br>

    <p>Dear {name.upper()},</p>

    <p>
        Attached is your e-Payslip for {payslip_period}.<br>
        Password: Your IC number without dashes.<br>
        Example: If your IC number is 980828-12-6128, the password is 980828126128.
    </p>

    <p>This is an auto-generated email. Please do not reply.</p>

    <p>
        Kind regards,<br><br>
        Payroll Unit
    </p>
</body>
</html>
""",
        subtype="html",
    )

    with open(
        pdf_path,
        "rb",
    ) as file_handle:

        message.add_attachment(
            file_handle.read(),
            maintype="application",
            subtype="pdf",
            filename=Path(pdf_path).name,
        )

    with smtplib.SMTP(
        SMTP_SERVER,
        SMTP_PORT,
        timeout=60,
    ) as server:

        server.ehlo()
        server.starttls()
        server.ehlo()

        server.login(
            SENDER_EMAIL,
            outlook_password,
        )

        server.send_message(message)


# ============================================================
# MANUAL REVIEW
# ============================================================


def move_page_to_manual_review(
    page_pdf,
    source_name,
    page_number,
    emp_id,
    reason,
):
    """Move a split page into manual review with a reason file."""

    source_stem = safe_filename_part(Path(source_name).stem)

    id_part = f"_emp_{emp_id}" if emp_id else ""

    review_filename = f"{source_stem}" f"_page_{page_number:03d}" f"{id_part}.pdf"

    destination = unique_destination(
        MANUAL_REVIEW_FOLDER,
        review_filename,
    )

    shutil.move(
        str(page_pdf),
        str(destination),
    )

    reason_file = destination.with_suffix(destination.suffix + ".reason.txt")

    with open(
        reason_file,
        "w",
        encoding="utf-8",
    ) as file_handle:

        file_handle.write(
            datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            + "\n"
            + f"Source: {source_name}\n"
            + f"Page: {page_number}\n"
            + f"Employee ID: {emp_id or 'Not detected'}\n"
            + reason
            + "\n"
        )

    print("Moved page to manual review: " f"{destination.name}")


# ============================================================
# PROCESS ONE PAGE
# ============================================================


def process_one_page(
    source_pdf,
    source_reader,
    source_hash,
    page_index,
    employees,
    log,
    outlook_password,
):
    """
    Split, identify, protect, send, and archive one payslip page.

    Returns:
        SUCCESS
        ALREADY_SENT
        MANUAL_REVIEW
        FAILED
    """

    page_number = page_index + 1

    page_key = f"{source_hash}:page:{page_number}"

    print("-" * 70)
    print(f"Page {page_number} " f"of {len(source_reader.pages)}")

    if page_already_sent(
        page_key,
        log,
    ):
        print("This page was already sent successfully. " "Skipping.")
        return "ALREADY_SENT"

    source_stem = safe_filename_part(source_pdf.stem)

    split_filename = f"{source_stem}" f"_page_{page_number:03d}.pdf"

    split_path = SPLIT_FOLDER / split_filename

    if split_path.exists():
        split_path.unlink()

    protected_path = None
    extraction_method = ""
    emp_id = ""
    email = ""

    try:
        create_single_page_pdf(
            source_reader=source_reader,
            page_index=page_index,
            output_pdf=split_path,
        )

        text, extraction_method = extract_pdf_text(split_path)
        payslip_period = extract_payslip_period(text)

        save_extracted_text(
            source_name=source_pdf.name,
            page_number=page_number,
            text=text,
            method=extraction_method,
        )

        emp_id = extract_employee_id(text)

        if not emp_id:
            reason = (
                "Employee ID could not be detected "
                "from this page. No email was sent."
            )

            append_log(
                log=log,
                emp_id="",
                email="",
                status="MANUAL_REVIEW",
                source_file=source_pdf.name,
                page_number=page_number,
                page_key=page_key,
                file_hash=source_hash,
                extraction_method=extraction_method,
                archived_file="",
                details=reason,
            )

            move_page_to_manual_review(
                page_pdf=split_path,
                source_name=source_pdf.name,
                page_number=page_number,
                emp_id="",
                reason=reason,
            )

            return "MANUAL_REVIEW"

        print(f"Detected Employee ID: {emp_id}")

        employee = find_employee(
            emp_id,
            employees,
        )

        if employee is None:
            reason = (
                f"Employee ID {emp_id} was detected, "
                "but there was no unique exact match "
                "in the employee Excel file. "
                "No email was sent."
            )

            append_log(
                log=log,
                emp_id=emp_id,
                email="",
                status="MANUAL_REVIEW",
                source_file=source_pdf.name,
                page_number=page_number,
                page_key=page_key,
                file_hash=source_hash,
                extraction_method=extraction_method,
                archived_file="",
                details=reason,
            )

            move_page_to_manual_review(
                page_pdf=split_path,
                source_name=source_pdf.name,
                page_number=page_number,
                emp_id=emp_id,
                reason=reason,
            )

            return "MANUAL_REVIEW"

        name = clean_excel_value(employee["Name"])

        email = clean_excel_value(employee["Email"])

        ic_number = clean_excel_value(employee["IC Number"])

        validation_errors = []

        if not name:
            validation_errors.append("employee name is empty")

        if not valid_email(email):
            validation_errors.append(f"email address is invalid: {email!r}")

        if not ic_number:
            validation_errors.append("IC number is empty")

        if validation_errors:
            reason = (
                f"Employee {emp_id}: "
                + "; ".join(validation_errors)
                + ". No email was sent."
            )

            append_log(
                log=log,
                emp_id=emp_id,
                email=email,
                status="MANUAL_REVIEW",
                source_file=source_pdf.name,
                page_number=page_number,
                page_key=page_key,
                file_hash=source_hash,
                extraction_method=extraction_method,
                archived_file="",
                details=reason,
            )

            move_page_to_manual_review(
                page_pdf=split_path,
                source_name=source_pdf.name,
                page_number=page_number,
                emp_id=emp_id,
                reason=reason,
            )

            return "MANUAL_REVIEW"

        protected_filename = f"{emp_id}" f"_{source_stem}" f".pdf"

        protected_path = PROTECTED_FOLDER / protected_filename

        if protected_path.exists():
            protected_path.unlink()

        protect_pdf(
            input_pdf=split_path,
            output_pdf=protected_path,
            password=ic_number,
        )

        send_email(
            email=email,
            emp_id=emp_id,
            name=name,
            pdf_path=protected_path,
            outlook_password=outlook_password,
            payslip_period=payslip_period,
        )

        archive_destination = unique_destination(
            ARCHIVE_FOLDER,
            protected_filename,
        )

        shutil.move(
            str(protected_path),
            str(archive_destination),
        )

        if split_path.exists():
            split_path.unlink()

        append_log(
            log=log,
            emp_id=emp_id,
            email=email,
            status="SUCCESS",
            source_file=source_pdf.name,
            page_number=page_number,
            page_key=page_key,
            file_hash=source_hash,
            extraction_method=extraction_method,
            archived_file=archive_destination.name,
            details=(
                "One-page payslip sent successfully "
                "and the protected PDF was archived."
            ),
        )

        print(f"Sent page {page_number}: " f"Employee ID {emp_id}, " f"email {email}")

        return "SUCCESS"

    except Exception as error:
        print(f"Page {page_number} failed: {error}")

        if protected_path and protected_path.exists():
            try:
                protected_path.unlink()
            except OSError:
                pass

        # A temporary split page is removed after an email/system failure.
        # The original multi-page source remains in input for retry.
        if split_path.exists():
            try:
                split_path.unlink()
            except OSError:
                pass

        append_log(
            log=log,
            emp_id=emp_id,
            email=email,
            status="FAILED",
            source_file=source_pdf.name,
            page_number=page_number,
            page_key=page_key,
            file_hash=source_hash,
            extraction_method=extraction_method,
            archived_file="",
            details=str(error),
        )

        write_error_log(
            f"Source file: {source_pdf}\n"
            f"Page: {page_number}\n"
            f"Employee ID: {emp_id}\n"
            f"Error: {error}"
        )

        return "FAILED"


# ============================================================
# MAIN PROCESS
# ============================================================


def process_payslips(outlook_password):
    """
    Process every multi-page PDF in the input folder.

    Each PDF page is treated as one employee payslip.
    """

    ensure_folders()

    employees = load_employee_data()
    log = load_log()

    pdf_files = sorted(
        path
        for path in PDF_FOLDER.iterdir()
        if (
            path.is_file()
            and path.suffix.lower() == ".pdf"
            and not path.name.lower().startswith("protected_")
        )
    )

    if not pdf_files:
        print("No PDF files found in the input folder.")
        return

    for source_pdf in pdf_files:
        print("\n" + "=" * 70)
        print(f"Processing source PDF: {source_pdf.name}")

        try:
            source_hash = calculate_file_hash(source_pdf)

            source_reader = PdfReader(str(source_pdf))

            total_pages = len(source_reader.pages)

            if total_pages == 0:
                raise ValueError("The source PDF contains no pages.")

            print(f"Total payslip pages: {total_pages}")

            page_results = []

            for page_index in range(total_pages):
                result = process_one_page(
                    source_pdf=source_pdf,
                    source_reader=source_reader,
                    source_hash=source_hash,
                    page_index=page_index,
                    employees=employees,
                    log=log,
                    outlook_password=outlook_password,
                )

                page_results.append(result)

            if every_page_sent(
                source_hash=source_hash,
                total_pages=total_pages,
                log=log,
            ):
                print("All pages have been sent successfully.")

                if DELETE_SOURCE_AFTER_ALL_PAGES_SENT and source_pdf.exists():
                    source_pdf.unlink()

                    print("Deleted completed original " "multi-page source PDF.")
            else:
                print(
                    "One or more pages were not sent. "
                    "The original multi-page PDF remains "
                    "in the input folder for retry."
                )

                print("Page results: " + ", ".join(page_results))

        except Exception as error:
            print(f"Source PDF failed: {error}")

            write_error_log(f"Source file: {source_pdf}\n" f"Error: {error}")


# ============================================================
# START
# ============================================================

if __name__ == "__main__":
    try:
        ensure_folders()

        outlook_password = os.getenv(
            "PAYROLL_EMAIL_PASSWORD",
            "",
        )

        if not outlook_password:
            outlook_password = getpass(f"Outlook password for " f"{SENDER_EMAIL}: ")

        if not outlook_password:
            raise ValueError("No Outlook password was provided.")

        # Stop immediately if Microsoft 365 rejects authentication.
        # This avoids OCR processing when email cannot be sent.
        test_smtp_login(outlook_password)

        process_payslips(outlook_password)

    except Exception as error:
        print("\nSYSTEM ERROR:")
        print(error)
        write_error_log(error)

    finally:
        input("\nPress ENTER to close...")
