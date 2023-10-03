"""
Author : Daniel Dekhtyar
Latest update : 1/10/2023
Version : 1.0

The code copies the coordinates of a specific area from an excel file,
to a table in Microsoft Word file called 'רכשי לב',or for short 'rahshal'
"""

"""
Text prints in revers bug solution : 
Before that only the first table was right and the rest were reversed.
I solved it by defining a global variable that keeps tack of the number of tables that is already modified.
If table_count == 0 means that we are still on the first table.
If it is bigger then 0 then we have passed the first table.
So to fix the bug it intentionally prints the values in reverse and the final product is as expected !
a.k.a backward logic
"""


import openpyxl
from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.enum.text import WD_ALIGN_PARAGRAPH
import time


start_time = time.time()

table_count = 0

def main():
    # Load the excel workbook
    excel_workbook = openpyxl.load_workbook(r"C:\Users\Daniel\Desktop\Iron Dome\Coordinates.xlsx")
    nz_xl = excel_workbook.active  # Open the active sheet; nz means נ.צ
    # Load the docx file
    rahshal = Document(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    xl_row = 1
    tables = rahshal.tables
    global table_count
    while xl_row < nz_xl.max_row:
        area_name, xl_row = find_area_in_xl(nz_xl, xl_row)
        if area_name != " ":
            table_index = find_area_in_rahshal(area_name, rahshal)
            # Check if None is returned, meaning that the area is not in the docx
            if table_index is not None:
                docx_table = tables[table_index]
                update_table_dimensions_in_rahshal(rahshal, nz_xl, xl_row, table_index)
                copy_coordinates_from_xl_to_rahshal(nz_xl, docx_table, xl_row)
                style_the_docx_table(docx_table)
                print("Successfully copied", area_name[::-1])
                table_count += 1

    rahshal.save(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    excel_workbook.close()
    print("")  # Just a white line
    print("All done and save successfully !")
    print(f"We've copied {table_count} tables from the excel file to rahshal")
    print(f"--- The code took {time.time() - start_time} seconds to run ---")


# 'TypeError: cannot unpack non-iterable NoneType object' solved by putting all the return values in to one tuple
# If the function didn't do it's work then return this one variable as None
# At the receiving end, check if the return value is None, else you can safely unpack the tuple and use it.


def find_area_in_xl(nz_xl, xl_row: int) -> (str, int):  # TESTED AND DONE !
    for row in range(xl_row + 1, nz_xl.max_row + 1):
        cell_value = nz_xl[f"A{row}"].value
        if cell_value is not None:
            if "קורדינטות" not in cell_value:
                return cell_value, row
    return " ", row


def find_area_in_rahshal(area_name: str, rahshal: Document) -> int or None:
    table_index = 0
    for table in rahshal.tables:
        if table_count == 0:
            if table.cell(1, 5).text == area_name:
                return int(table_index)
        else:
            if table.cell(1, 0).text == area_name:
                return int(table_index)
        table_index += 1
    # If area is not in rahshal, an error message will appear
    print("----------------------------------------------")
    print(f"The area {area_name[::-1]} is not in the docx")
    print("----------------------------------------------")
    return None


def update_table_dimensions_in_rahshal(rahshal: Document, nz_xl, xl_row : int, table_index: int) -> None:
    tables = rahshal.tables
    docx_table = tables[table_index]
    rows_old_table = len(docx_table.rows)
    rows_new_table = xl_table_dimensions(nz_xl, xl_row) + 3
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
    # Finds the number of rows in the table
    int(xl_row)
    first_row_of_table = None
    last_row_of_table = None
    # I added +1 to max_row because otherwise it will go upto the one to last but not the last row
    max_row = nz_xl.max_row + 1

    for row in range(int(xl_row), max_row):
        cell_in_column_B = nz_xl.cell(row, 2)  # 2 corresponds to column B
        if cell_in_column_B.value is not None:
            first_row_of_table = row
            break

    for row in range(first_row_of_table, max_row):
        cell_in_column_B = nz_xl.cell(row, 2)  # 2 corresponds to column B
        if cell_in_column_B.value is None:
            last_row_of_table = row
            break
        elif row == max_row:
            last_row_of_table = max_row

    number_of_rows_in_table = last_row_of_table - first_row_of_table
    return number_of_rows_in_table

# Copies the content of the table from excel to docx
def copy_coordinates_from_xl_to_rahshal(nz_xl, docx_table: Document, xl_row: int) -> None:
    # Start at row 2 because we need to leave space for the 2 default rows
    for row in range(5, len(docx_table.rows)):
        if table_count == 0:
            for column in range(1, 7):  # Columns 1 to 7 in the excel
                cell = nz_xl.cell((int(xl_row) + row) - 1, column + 1) 
                # 'row + 2' because don't need to copy the first 2 rows
                if cell.value is not None:
                    docx_table.cell(row, column - 1).text = str(cell.value)
                    # 'column - 1' because in docx it start from index 0 and in excel it starts from index 1
        else:
            for column in range(1, 7):  # Columns 1 to 7 in the excel
                cell = nz_xl.cell((int(xl_row) + row) - 1, column + 1) 
                # 'row + 2' because don't need to copy the first 2 rows
                if cell.value is not None:
                    docx_table.cell(row, abs(column - 6)).text = str(cell.value)
                    # 'abs(column - 6)' explanation in the doc in the start of the document

def style_the_docx_table(docx_table: Document) -> None:
    for row in docx_table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.highlight_color = WD_COLOR_INDEX.YELLOW
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run.font.name = "Calibri Light"

if __name__ == "__main__":
    main()
