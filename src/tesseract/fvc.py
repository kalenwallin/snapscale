# Import required libraries
import pytesseract
from openpyxl import Workbook
import cv2

# Check src directory for packages
import sys

sys.path.append("./src")

# Import local libraries
import src.tesseract.img_manipulation as img_manipulation

# Your image folder and file
image_dir = "documents/"
image_file = "fvc1.jpg"
image_path = image_dir + image_file
image_path = "documents/fvc/505344-scan.jpg"

# Load the image from disk
img = cv2.imread(image_path)

# Convert to grayscale
gray = img_manipulation.get_grayscale(img)

# Thresholding
_, img_bin = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

# Inverting the image
img_bin = 255 - img_bin

# Apply OCR to the image
text = pytesseract.image_to_string(img_bin)

# Print extracted text
print(text)
print("\n====\n")

# Process the text as required and write to Excel
# Create a workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Let's assume the extracted text is a list of values
values = text.split()

for idx, value in enumerate(values, start=1):
    # This assumes you want each value in a separate cell in the first column
    ws.cell(row=idx, column=1, value=value)
    print(str(idx) + " : " + value)

# Save the workbook
wb.save("output/fvc_output_scan.xlsx")
