import datetime
from openpyxl import Workbook
from openpyxl.styles import NamedStyle, Font, Alignment, Border, Side

from src.classes import generate_mock_ticket


def write_header(workbook):
    """
    Writes the header section of the workbook including contact info, location, and headers.
    Also, sets the column widths and row heights of the worksheet.

    Parameters:
    - workbook: The workbook object to modify.

    Returns:
    - The modified workbook object.
    """
    worksheet = workbook.active

    # Column Width
    # The Excel column width (10.29) divided by openpyxl Excel column width (9.57) to get the actual Excel column width
    worksheet.column_dimensions["A"].width = 10.29 * (10.29 / 9.57)
    worksheet.column_dimensions["B"].width = 8 * (8 / 7.29)
    worksheet.column_dimensions["C"].width = 9.43 * (9.43 / 8.71)
    worksheet.column_dimensions["D"].width = 8 * (8 / 7.29)
    worksheet.column_dimensions["E"].width = 5.86 * (5.86 / 5.14)
    worksheet.column_dimensions["F"].width = 6.29 * (6.29 / 5.57)
    worksheet.column_dimensions["G"].width = 9.71 * (9.71 / 9)
    worksheet.column_dimensions["H"].width = 11.86 * (11.86 / 11.14)
    worksheet.column_dimensions["I"].width = 10.86 * (10.86 / 10.14)
    worksheet.column_dimensions["J"].width = 10.29 * (10.29 / 9.57)
    worksheet.column_dimensions["K"].width = 6.57 * (6.57 / 5.86)

    # Row Height
    for row in range(4, 11):
        worksheet.row_dimensions[row].height = 12.75

    # Contact Info
    contact_style = NamedStyle(name="contact")
    contact_style.font = Font(name="Arial", size=24)
    contact_style.alignment = Alignment(horizontal="left", wrapText=False)
    workbook.add_named_style(contact_style)
    worksheet["A1"] = "Jeff Wallin"
    worksheet["A1"].style = "contact"
    worksheet["A2"] = "PO Box 240"
    worksheet["A2"].style = "contact"
    worksheet["A3"] = "Imperial NE 69033"
    worksheet["A3"].style = "contact"

    ## Borders
    border_side = Side(border_style="thin", color="000000")
    worksheet["A1"].border = Border(left=border_side, top=border_side)
    worksheet["B1"].border = Border(top=border_side)
    worksheet["C1"].border = Border(right=border_side, top=border_side)
    worksheet["A2"].border = Border(left=border_side)
    worksheet["C2"].border = Border(right=border_side)
    worksheet["A3"].border = Border(left=border_side, bottom=border_side)
    worksheet["B3"].border = Border(bottom=border_side)
    worksheet["C3"].border = Border(right=border_side, bottom=border_side)

    # Location
    worksheet["C6"] = "Sold To:"
    worksheet["C6"].font = Font(name="Arial", size=8)
    worksheet["C6"].alignment = Alignment(horizontal="center")
    worksheet["H6"] = "Field"
    worksheet["H6"].font = Font(name="Arial", size=10)
    worksheet["H6"].alignment = Alignment(horizontal="center")

    # Headers
    header_style = NamedStyle(name="header")
    header_style.font = Font(name="Arial", size=10, bold=True)
    header_style.alignment = Alignment(horizontal="center")
    workbook.add_named_style(header_style)
    worksheet["A10"] = "Date"
    worksheet["A10"].style = "header"
    worksheet["B10"] = "Ticket"
    worksheet["B10"].style = "header"
    worksheet["C10"] = "Gross"
    worksheet["C10"].style = "header"
    worksheet["D10"] = "Tare"
    worksheet["D10"].style = "header"
    worksheet["E10"] = "MO"
    worksheet["E10"].style = "header"
    worksheet["F10"] = "TW"
    worksheet["F10"].style = "header"
    worksheet["G10"] = "Net"
    worksheet["G10"].style = "header"
    worksheet["H10"] = "Wet Bu"
    worksheet["H10"].style = "header"
    worksheet["I10"] = "Dry Bu"
    worksheet["I10"].style = "header"
    worksheet["J10"] = "Driver"
    worksheet["J10"].style = "header"
    worksheet["K10"] = "Truck"
    worksheet["K10"].style = "header"

    return workbook


def write_totals(workbook, numTickets):
    """
    Writes the totals to the worksheet.

    Args:
        workbook (Workbook): The workbook to write to.
        numTickets (int): The number of corn tickets processed.

    Returns:
        Workbook: The updated workbook.
    """
    worksheet = workbook.active

    # Total style definition
    total_style = NamedStyle(name="total")
    total_style.font = Font(name="Arial", size=12, bold=True)
    center_aligned_text = Alignment(horizontal="center")
    total_style.alignment = center_aligned_text
    workbook.add_named_style(total_style)

    # The insertion row should be numTickets + 13 (1 rows below the last ticket).
    row = numTickets + 13

    worksheet["A" + str(row)] = "Totals"
    worksheet["A" + str(row)].style = total_style

    for letter in "CDGHI":
        start_cell = letter + "12"  # Data starts at row 12
        end_cell = letter + str(numTickets + 11)  # +11 because data starts at row 12
        formula = f"=SUM({start_cell}:{end_cell})"
        cell = worksheet[letter + str(row)]
        cell.value = formula
        cell.style = total_style
        if letter == "H" or letter == "I":
            cell.number_format = "##,##0.00"  # Include comma for thousands

    row += 2
    worksheet["A" + str(row)] = "Acres"
    worksheet["A" + str(row)].alignment = center_aligned_text
    row += 1
    worksheet["A" + str(row)] = "Yield"
    worksheet["A" + str(row)].alignment = center_aligned_text

    return workbook


# Testing
if __name__ == "__main__":
    import sys

    if sys.argv[1] == "write_header":
        workbook = Workbook()
        workbook = write_header(workbook)
        workbook.save("output/test_write_header.xlsx")
    elif sys.argv[1] == "write_sheet":
        workbook = Workbook()
        workbook = write_header(workbook)

        # Write mock tickets
        worksheet = workbook.active

        # Write sale location
        worksheet["C7"] = "Baney Feeders"
        worksheet["C7"].font = Font(name="Arial", size=11)
        worksheet["C7"].alignment = Alignment(horizontal="center")

        # Write tickets
        import string

        ticket_style = NamedStyle(name="ticket")
        ticket_style.font = Font(name="Arial", size=10)
        ticket_style.alignment = Alignment(horizontal="center")
        workbook.add_named_style(ticket_style)

        tickets = [generate_mock_ticket() for _ in range(2)]  # List of tickets

        attributes_order = [
            "date",
            "ticket",
            "gross",
            "tare",
            "mo",
            "tw",
            "net",
            "wetBu",
            "dryBu",
            "driver",
            "truck",
        ]

        from openpyxl.styles import Alignment

        # Center Alignment
        center_aligned_text = Alignment(horizontal="center")

        for i, ticket in enumerate(tickets, start=12):
            letters = iter(string.ascii_uppercase)
            attribute_actions = {
                "wetBu": lambda val: round(val, 2),
                "dryBu": lambda val: round(val, 2),
                "ticket": round,
                "gross": round,
                "tare": round,
                "net": round,
                "mo": lambda val: round(val, 1),
                "tw": lambda val: round(val, 1),
            }

            for letter, attr in zip(letters, attributes_order):
                attr_value = getattr(ticket, attr)
                cell_value = attribute_actions.get(attr, lambda x: x)(attr_value)
                cell = worksheet[letter + str(i)]
                cell.value = cell_value
                cell.style = ticket_style
                if attr == "wetBu" or attr == "dryBu":
                    cell.alignment = center_aligned_text
                    cell.number_format = "#,##0.00"  # Include comma for thousands
                if attr == "date":
                    cell.number_format = "mm/dd/yyyy"

        workbook = write_totals(workbook, 2)

        # Set page orientation to landscape
        workbook.active.page_setup.orientation = "landscape"

        # Save the workbook
        workbook.save("output/test_write_sheet.xlsx")
    else:
        print("Error: No such function exists")
