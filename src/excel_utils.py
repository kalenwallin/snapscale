from openpyxl import Workbook

def write_header(worksheet):
    ws.cell(row=1, column=1, value="Jeff Wallin")
    ws.cell(row=2, column=1, value="PO Box 240")
    ws.cell(row=3, column=1, value="Imperial NE 69033")
    
    ws.cell(row=6, column=3, value="Sold To:")
    ws.cell(row=6, column=8, value="Field")
    
    ws.cell(row=10, column=1, value="Date")
    ws.cell(row=10, column=2, value="Ticket")
    ws.cell(row=10, column=3, value="Gross")
    ws.cell(row=10, column=4, value="Tare")
    ws.cell(row=10, column=5, value="MO")
    ws.cell(row=10, column=6, value="TW")
    ws.cell(row=10, column=7, value="MO")
    return worksheet

# Testing
if __name__ == "__main__":
    import sys
    if sys.argv[1] == "write_header":
        # Create a workbook and select the active worksheet
        wb = Workbook()
        ws = wb.active
        write_header(ws)
    else:
        print("Error: No such function exists")