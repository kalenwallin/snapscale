# Import required libraries
import pytesseract
from openpyxl import Workbook
import cv2

# Your image folder and file
image_dir = "../documents/"
image_file = "baney1.jpg"
image_path = image_dir + image_file

# Load the image from disk
img = cv2.imread(image_path)

# Convert to grayscale
gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

# Thresholding 
_, img_bin = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

# Inverting the image
img_bin = 255-img_bin

# Apply OCR to the image
text = pytesseract.image_to_string(img_bin)

# Print extracted text
print(text)

# Process the text as required and write to Excel
# Create a workbook and select the active worksheet
wb = Workbook()
ws = wb.active

# Let's assume the extracted text is a list of values
values = text.split()

for idx, value in enumerate(values, start=1):
    # This assumes you want each value in a separate cell in the first column
    ws.cell(row=idx, column=1, value=value)

# Save the workbook
wb.save("../output/baney_output.xlsx")
