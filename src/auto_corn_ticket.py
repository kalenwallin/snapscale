# Import required libraries
import pytesseract
from openpyxl import Workbook
import cv2

# Check src directory for packages
import sys
sys.path.append("./src")

# Import local libraries
import img_manipulation

# TODO: Write the beginning of the Excel sheet
# TODO: Jeff Wallin, PO Box 240, Imperial NE 69033
# TODO: Sold To: {Feedlot} Field: {Field #}
# TODO: Feedlot and Field # can be added from User Interface
# TODO: Headers: Date, Ticket #, Gross, Tare, MO, TW, Net, Wet Bu, Dry Bu, Driver, Truck

# Your image folder from argument 1
image_dir = sys.argv[1]

# TODO: Iterate through image_dir's jpg files
# TODO: For each jpg, process it as a line in an Excel sheet

# Your image file
image_file = "fvc1.jpg"
image_path = image_dir + image_file

# Load the image from disk
img = cv2.imread(image_path)

# Convert to grayscale
gray = img_manipulation.get_grayscale(img)

# Thresholding 
_, img_bin = cv2.threshold(gray, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

# Inverting the image
img_bin = 255-img_bin

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
    
# TODO: Write the Totals line: Totals, ___, Gross, Tare, ___, ___, Net, Wet Bu, Dry Bu, ____, ____
# TODO: Write Acres Yield

# Save the workbook
wb.save("output/fvc_output.xlsx")
