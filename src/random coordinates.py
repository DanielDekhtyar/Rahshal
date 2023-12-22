import openpyxl as xl
from random import randint

"""
Create a random coordinate sheet in excel
It just helped me do the creation of the random coordinates faster
Don't have a role in the actual program
"""


def main():
    """
    The main function loads an Excel workbook, creates random areas in the sheet, saves the workbook,
    and prints a completion message.
    """
    wb = xl.load_workbook(r"C:\Users\Daniel\Desktop\Tests\Coordinates.xlsx")
    sheet = wb.active

    new_row = 0
    for _ in range(10):
        # Create random area length between 4 and 50 and return the last row to last_row
        new_row = create_random_area(sheet, new_row)
        new_row = new_row + 5  # Add 5 to last_row to get the next area

    wb.save(r"C:\Users\Daniel\Desktop\Tests\Coordinates.xlsx")
    print("All done and saved!")


def create_random_area(sheet, start_raw):
    """
    The function creates a random area of numbers in a given sheet starting from a specified row.

    Args:
    sheet: The "sheet" parameter is the Excel sheet object where the random area will be created. It
    is assumed that the "sheet" object has already been defined and passed as an argument to the
    function.
    start_raw: The parameter "start_raw" represents the starting row number from which the random area
    will be created.

    Returns:
    The row number of the last cell that was modified.
    """
    for row in range(start_raw + 4, start_raw + randint(4, 50)):
        for column in range(3, 9):
            cell = sheet.cell(row, column)
            cell.value = randint(1000000, 9999999)
    return row


main()
