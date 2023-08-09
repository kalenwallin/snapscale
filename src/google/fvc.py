# Import required libraries
from google.cloud import vision
from openpyxl import Workbook

# Check src directory for packages
import sys

sys.path.append("./src")

# Your image path
image_path = "documents/fvc/505344-scan.jpg"

# Instantiate a Vision client
client = vision.ImageAnnotatorClient()

# Read the image file
with open(image_path, "rb") as image_file:
    content = image_file.read()

image = vision.Image(content=content)

response = client.text_detection(image=image)
texts = response.text_annotations
print("Texts:")

for text in texts:
    print(f'\n"{text.description}"')

    vertices = [f"({vertex.x},{vertex.y})" for vertex in text.bounding_poly.vertices]

    print("bounds: {}".format(",".join(vertices)))

if response.error.message:
    raise Exception(
        "{}\nFor more info on error messages, check: "
        "https://cloud.google.com/apis/design/errors".format(response.error.message)
    )

# Process the text as required and write to Excel
# Create a workbook and select the active worksheet
# workbook = Workbook()
# worksheet = workbook.active


# # Save the workbook
# workbook.save("output/fvc_output_scan.xlsx")
