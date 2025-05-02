import cv2
import pandas as pd
import numpy as np
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from PIL import Image
import os
import time
import subprocess
from pdf2image import convert_from_path
import qrcode
from reportlab.lib.units import mm
import platform
import sys

# ---- CONFIG ----
CSV_PATH = 'GOSIM_AI_PARIS_Attendees_1217254842023_20250501_080105_527.csv'
SUB_CATEGORY = 'SPEAKER' # Change to needed sub-category (please use the 4 sub-categories in PNG folder)
BACKGROUND_IMAGE_PATH = f'./badge_template/PNG/{SUB_CATEGORY}.png'
PDF_OUTPUT_DIR = f'./generated_pdfs/{SUB_CATEGORY}'
SCAN_INTERVAL = 0.1
CAMERA_INDEX = 0
PRINT_COMMAND = 'lp'

os.makedirs(PDF_OUTPUT_DIR, exist_ok=True)

# ---- Load CSV ----
df = pd.read_csv(CSV_PATH)
df['id'] = df['Order ID'].astype(str)

# ---- Alarms ----
def alert_beep():
    system = platform.system()
    try:
        if system == "Windows":
            import winsound
            winsound.MessageBeep()
        elif system == "Darwin":
            # macOS: use built-in system bell
            os.system('say "error"')  # read aloud "error"
        else:
            # Linux: system bell
            print('\a')  # ASCII Bell
    except Exception as e:
        print(f"[Sound alert failed] {e}", file=sys.stderr)

# ---- PDF Generation ----
def generate_pdf(entry_data, output_path):
    bg_image = Image.open(BACKGROUND_IMAGE_PATH)
    print(f"Background image size (pixels): width={bg_image.width}, height={bg_image.height}")

    # width, height = bg_image.size
    # c = canvas.Canvas(output_path, pagesize=(width, height))
    page_width = 96 * mm
    page_height = 278 * mm
    c = canvas.Canvas(output_path, pagesize=(page_width, page_height))

    # Draw background and drag image to fit
    c.drawImage(ImageReader(bg_image), 0, 0, width=page_width, height=page_height)

    # Prepare name
    first_name = entry_data['Attendee first name'].strip().title()
    last_name = entry_data['Attendee last name'].strip().upper()  # Uppercase last name
    full_name_lines = [first_name, last_name]

    # Font settings
    font_name = "Helvetica-Bold"
    font_size = 28
    c.setFont(font_name, font_size)

    # Calculate vertical start position: under "GOSIM AI PARIS 2025"
    start_y = page_height - 65*mm  # parameter adjusted

    # Draw each line centered
    for i, line in enumerate(full_name_lines):
        text_width = c.stringWidth(line, font_name, font_size)
        x = (page_width - text_width) / 2
        y = start_y - i * (font_size + 10)
        c.drawString(x, y, line)

    # ---- Generate new QR code containing Order id under the name ----
    qr_id = str(entry_data['id'])[:11]
    qr_img = qrcode.make(qr_id)

    qr_size = 22*mm  # mm
    qr_img = qr_img.resize((int(qr_size), int(qr_size)), Image.LANCZOS)

    # Calculate position: centered below last name
    qr_x = (page_width - qr_size) / 2
    qr_y = start_y - len(full_name_lines) * (font_size + 10) - qr_size + 3*mm

    c.drawImage(ImageReader(qr_img), qr_x, qr_y, width=qr_size, height=qr_size)

    page_width_pt, page_height_pt = c._pagesize  # unit: points
    # Convert points to mm
    page_width_mm = page_width_pt / mm
    page_height_mm = page_height_pt / mm
    print(f"Generated PDF size: {page_width_mm:.2f} mm × {page_height_mm:.2f} mm")

    c.save()


# ---- Printing ----
def print_pdf(file_path):
    try:
        if os.name == 'nt':
            os.startfile(file_path, "print")
        else:
            subprocess.run([PRINT_COMMAND, file_path])
    except Exception as e:
        print(f"Error printing: {e}")

# ---- Webcam QR Detection ----
def scan_qr_and_generate():
    cap = cv2.VideoCapture(CAMERA_INDEX)
    detector = cv2.QRCodeDetector()
    scanned_ids = set()
    output_pdf = None
    print("Starting webcam scan... Press ctrl+C to quit.")

    while True:
        ret, frame = cap.read()
        if not ret:
            continue

        data, bbox, _ = detector.detectAndDecode(frame)
        if data:
            qr_id = data[:11]
            if qr_id not in scanned_ids:
                print(f"Detected QR Code ID: {qr_id}")
                scanned_ids.add(qr_id)

                matched = df[df['id'] == qr_id]

                if not matched.empty:
                    entry = matched.iloc[0]
                    try:
                        output_pdf = os.path.join(PDF_OUTPUT_DIR, f"{qr_id}.pdf")
                        generate_pdf(entry, output_pdf)
                        print(f"PDF generated: {output_pdf}")
                        print_pdf(output_pdf)
                        print(f"Printed: {output_pdf}")
                    except Exception as e:
                        print(f"Error during generating or printing PDF: {e}")
                        alert_beep()
                else:
                    print(f"No matching entry found for ID: {qr_id}")
                    alert_beep()

        cv2.imshow('QR Scanner', frame)
        if cv2.waitKey(1) == 27:
            break
        time.sleep(SCAN_INTERVAL)

    cap.release()
    cv2.destroyAllWindows()
    return output_pdf

# ---- Test QR in Sample PDF ----
def test_sample_pdf():
    print("Running test: Extracting QR from sample PDF...")

    images = convert_from_path('./1265928669729-12019541093-ticket.pdf')
    if not images:
        print("Failed to convert PDF.")
        return

    image = cv2.cvtColor(np.array(images[0]), cv2.COLOR_RGB2BGR)
    detector = cv2.QRCodeDetector()
    data, bbox, _ = detector.detectAndDecode(image)

    if not data:
        print("No QR code found.")
        return
    
    qr_id = data[:11]
    print(f"Detected QR code: {qr_id}")
    matched = df[df['id'] == qr_id]
    if not matched.empty:
        entry = matched.iloc[0]
        output_pdf = os.path.join(PDF_OUTPUT_DIR, f"test_{data}.pdf")
        generate_pdf(entry, output_pdf)
        print(f"✅ Test PDF generated: {output_pdf}")
        return output_pdf
    else:
        print(f"⚠️ No match found for ID: {qr_id}")
        return None

# ---- Entry point ----
if __name__ == "__main__":
    output_pdf = scan_qr_and_generate()  # Uncomment to use webcam
    # test_sample_pdf()         # Comment this out if not testing sample PDF
