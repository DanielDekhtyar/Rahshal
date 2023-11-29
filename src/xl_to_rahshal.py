"""
Author : Daniel Dekhtyar
Latest update : 29/11/2023
Version : 1.0.3

The code copies the coordinates of a specific area from an excel file,
to a table in Microsoft Word file called 'רכשי לב',or for short 'rahshal'

Changelog :
1.0.3 (29-11-2023)
>> sys library not imports as it is not needed
>> The success or failure message are now printed in color (green or red respectfully)
>> Error messages are printed in bold
>> Comments and documentation made more clear and understandable
>> Documentation added to every function
>> Code reformated with pylint

1.0.2 (3-11-2023)
>> Bug fix :
    - first_row_of_table and last_row_of_table initialization changed from None to '0'
    - comments added

1.0.1 (4-10-2023)
>> Minor code readability improvement

1.0.0 (2-10-2023)
>> First fully working version of the code.
>> find_area_in_rahshal() reimplemented to fit the new docx format
"""

"""
Text prints in revers bug solution : 
Before that solution, only the first table printed the right way and the rest were reversed.
I solved it by defining a global variable (table_count) that keeps track of the number of tables that is already modified.
If table_count == 0 means that we are still on the first table.
If it is bigger then 0 then we have passed the first table.
So to fix the bug it intentionally prints the values in reverse and the final product is as expected !
a.k.a backward logic

table_count is also used to report on the total number of tables copied at the end of the program.
"""


import time
import openpyxl
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.text import WD_ALIGN_PARAGRAPH
from termcolor import colored


start_time = time.time()

TABLE_COUNT: int = 0


def main():
    """
    The main function loads an Excel workbook and a Word document, searches for specific areas in the
    Excel file, finds corresponding tables in the Word document, copies coordinates from the Excel file
    to the Word document, styles the tables in the Word document, and saves the updated Word document.
    """
    # Load the excel workbook
    excel_workbook = openpyxl.load_workbook(
        r"C:\Users\Daniel\Desktop\Iron Dome\Coordinates.xlsx"
    )
    nz_xl = excel_workbook.active  # Open the active sheet; nz means נ.צ
    # Load the docx file
    rahshal = Document(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    xl_row: int = 1  # Stores the row number of the last area found in excel
    tables = rahshal.tables  # Count how many tables copied
    global TABLE_COUNT
    while xl_row < nz_xl.max_row:
        area_name, xl_row = find_area_in_xl(nz_xl, xl_row)
        if area_name != " ":
            table_index: int = find_area_in_rahshal(area_name, rahshal)
            # Check if None is returned, meaning that the area is not in the docx
            if table_index is not None:
                docx_table = tables[table_index]
                update_table_dimensions_in_rahshal(rahshal, nz_xl, xl_row, table_index)
                copy_coordinates_from_xl_to_rahshal(nz_xl, docx_table, xl_row)
                style_the_docx_table(docx_table)
                print(colored(f"Successfully copied {area_name[::-1]}", "green"))
                TABLE_COUNT += 1

    rahshal.save(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    excel_workbook.close()
    print("")  # Just a white line
    print("All done and save successfully !")
    print(f"We've copied {TABLE_COUNT} tables from the excel file to rahshal")
    print(f"--- The code took {time.time() - start_time} seconds to run ---")
    print("")  # Just a white line


# 'TypeError: cannot unpack non-iterable NoneType object' solved by putting all the return values in to one tuple
# If the function didn't do it's work then return this one variable as None
# At the receiving end, check if the return value is None, else you can safely unpack the tuple and use it.


def find_area_in_xl(nz_xl, xl_row: int) -> [str, int]:  # TESTED AND DONE !
    """
    The function `find_area_in_xl` finds the first non-empty cell value in column A of an Excel sheet
    starting from a given row.

    :param nz_xl: The parameter `nz_xl` is expected to be an object representing an Excel file or
    worksheet. It is used to access the cells in the Excel file
    :param xl_row: The parameter `xl_row` is an integer that represents the current row number in an
    Excel spreadsheet
    :type xl_row: int
    :return: The function `find_area_in_xl` returns a tuple containing a string and an integer. The
    string represents the cell value found in the Excel sheet, and the integer represents the row number
    where the cell value was found.
    """
    for row in range(xl_row + 1, nz_xl.max_row + 1):
        cell_value = nz_xl[f"A{row}"].value
        if cell_value is not None and "קורדינטות" not in cell_value:
            return cell_value, row
    return " ", row


def find_area_in_rahshal(area_name: str, rahshal: Document) -> [int or None]:
    """
    The function `find_area_in_rahshal` takes an area name and a document object as input, and returns
    the index of the table containing the area name in the document, or None if the area is not found.

    :param area_name: The area name is a string that represents the name of the area you want to find in
    the "rahshal" document
    :type area_name: str
    :param rahshal: The parameter "rahshal" is of type "Document", which suggests that it is a document
    object, possibly representing a Microsoft Word document
    :type rahshal: Document
    :return: The function `find_area_in_rahshal` returns an integer value representing the index of the
    table in the `rahshal` document that contains the specified `area_name`. If the `area_name` is not
    found in the document, the function returns `None`.
    """
    for table_index, table in enumerate(rahshal.tables):
        if TABLE_COUNT == 0:
            if table.cell(1, 5).text == area_name:
                return int(table_index)
        elif table.cell(1, 0).text == area_name:
            return int(table_index)
        table_index += 1
    # If area is not in the rahshal docx, this error message will appear
    print(colored("---------------------------------", "red", attrs=["bold"]))
    print(
        colored(f"The area {area_name[::-1]} is not in the docx", "red", attrs=["bold"])
    )
    print(colored("---------------------------------", "red", attrs=["bold"]))
    return None


def update_table_dimensions_in_rahshal(
    rahshal: Document, nz_xl, xl_row: int, table_index: int
) -> None:
    """
    The function `update_table_dimensions_in_rahshal` updates the dimensions of a table in a document by
    adding or removing rows based on the specified number of rows from an Excel file.

    :param rahshal: The `rahshal` parameter is a `Document` object, which represents a Word document
    :type rahshal: Document
    :param nz_xl: The parameter `nz_xl` is not defined in the given code snippet. It seems to be a
    missing variable or function that is required to determine the number of rows in the Excel table.
    Please provide more information or update the code snippet with the definition of `nz_xl` so that I
    :param xl_row: The `xl_row` parameter represents the row number in the Excel file
    :type xl_row: int
    :param table_index: The parameter `table_index` is the index of the table in the `rahshal` document
    that you want to update the dimensions for. It is used to access the specific table in the `tables`
    list
    :type table_index: int
    """
    tables = rahshal.tables
    docx_table = tables[table_index]
    rows_old_table: int = len(docx_table.rows)
    rows_new_table: int = xl_table_dimensions(nz_xl, xl_row) + 3
    # Add +3 to rows_new_table because 3 rows added to the docx_table to have the new format

    if rows_old_table == rows_new_table:
        pass
    elif rows_old_table < rows_new_table:
        rows_to_add = rows_new_table - rows_old_table
        for _ in range(rows_to_add):
            docx_table.add_row()
    elif rows_old_table > rows_new_table:
        rows_to_remove = rows_old_table - rows_new_table
        for _ in range(rows_to_remove):
            row = docx_table.rows[len(docx_table.rows) - 1]
            table_element = docx_table._tbl
            row_element = row._tr
            table_element.remove(row_element)


def xl_table_dimensions(nz_xl, xl_row: int) -> int:  # TESTED AND DONE !
    """
    The `xl_table_dimensions` function takes in an Excel worksheet (`nz_xl`) and a starting row
    (`xl_row`), and returns the number of rows in a table starting from the given row.

    :param nz_xl: A variable representing the Excel file or worksheet that contains the table
    :param xl_row: The `xl_row` parameter is the starting row from which you want to find the table
    dimensions. It is an integer value representing the row number
    :type xl_row: int
    :return: The function `xl_table_dimensions` returns the number of rows in a table.
    """
    # Finds the number of rows in the table
    first_row_of_table: int = 0
    last_row_of_table: int = 0
    # I added +1 to max_row because otherwise it will go upto the one to last but not the last row
    max_row = nz_xl.max_row + 1

    for row in range(xl_row, max_row):
        cell_in_column_B = nz_xl.cell(row, 2)  # 2 corresponds to column B
        if cell_in_column_B.value is not None:
            first_row_of_table = row
            break

    for row in range(first_row_of_table, max_row):
        cell_in_column_B = nz_xl.cell(row, 2)  # 2 corresponds to column B
        if cell_in_column_B.value is None:
            last_row_of_table = row
            break
        if row == max_row:
            last_row_of_table = max_row

    return last_row_of_table - first_row_of_table


def copy_coordinates_from_xl_to_rahshal(
    nz_xl, docx_table: Document, xl_row: int
) -> None:
    """
    The function `copy_coordinates_from_xl_to_rahshal` copies coordinates from an Excel file to a
    specified table in a Word document.

    :param nz_xl: The variable `nz_xl` is likely referring to an Excel workbook or worksheet object that
    contains the coordinates data. It is used to access specific cells in the Excel sheet
    :param docx_table: The parameter `docx_table` is of type `Document`, which is likely referring to a
    document object in a word processing software such as Microsoft Word. This object represents a
    document that contains tables
    :type docx_table: Document
    :param xl_row: The `xl_row` parameter represents the row number in the Excel sheet from which the
    coordinates will be copied
    :type xl_row: int
    """
    # Start at row 5 because we need to leave space for the 5 default rows
    for row in range(5, len(docx_table.rows)):
        if TABLE_COUNT == 0:
            for column in range(1, 7):  # Columns 1 to 7 in the excel
                cell = nz_xl.cell((xl_row + row) - 1, column + 1)
                # 'row + 2' because don't need to copy the first 2 rows
                if cell.value is not None:
                    docx_table.cell(row, column - 1).text = str(cell.value)
                    # 'column - 1' because in docx it start from index 0 and in excel it starts from index 1
        else:
            for column in range(1, 7):  # Columns 1 to 7 in the excel
                cell = nz_xl.cell((xl_row + row) - 1, column + 1)
                # 'row + 2' because don't need to copy the first 2 rows
                if cell.value is not None:
                    docx_table.cell(row, abs(column - 6)).text = str(cell.value)
                    # 'abs(column - 6)' explanation in the doc in the start of the document


def style_the_docx_table(docx_table: Document) -> None:
    """
    The function `style_the_docx_table` applies specific formatting to a table in a Word document.

    :param docx_table: The parameter `docx_table` is of type `Document`, which represents a Word
    document in the python-docx library. It is assumed that `docx_table` contains a table that needs to
    be styled
    :type docx_table: Document
    """
    for row in docx_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run.font.name = "Calibri Light"


if __name__ == "__main__":
    main()
