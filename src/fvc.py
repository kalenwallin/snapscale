# Import required libraries
from PIL import Image
import pytesseract
from openpyxl import Workbook

# Your image folder and file
image_dir = "../documents/"
image_file = "fvc1.jpg"
image_path = image_dir + image_file

# Open and process the image
img = Image.open(image_path)

# Apply OCR to the image
text = pytesseract.image_to_string(img)

# Print the resulting text
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
wb.save("../output/fvc_output.xlsx")
