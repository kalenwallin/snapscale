# Import required libraries
# Check src directory for packages
from openpyxl import Workbook

from google.cloud import vision
from src.classes import process_image_as_corn_ticket
from src.excel_utils import write_sheet

# Your image path
image_path = "documents/fvc/505344-scan.jpg"

# Instantiate a Vision client
client = vision.ImageAnnotatorClient()

tickets = []
corn_ticket = process_image_as_corn_ticket(client, image_path)
tickets.append(corn_ticket)

# Process the text as required and write to Excel
# Create a workbook and select the active worksheet
wb = Workbook()
ws = wb.active

write_sheet(ws, tickets, "FVC", "")

# Save the workbook
wb.save("output/fvc_single.xlsx")
