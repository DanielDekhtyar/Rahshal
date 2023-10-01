"""
Author : Daniel Dekhtyar
Latest update : 2/10/2023

The code adds 3 rows at the top of the table in formats it as required
"""

"""
Text prints in revers bug solution : 
Before that only the first table was right and the rest were reversed.
I solved it by defining a global variable that keeps tack of the number of tables that is already modified.
If table_count = 0 means that we are still on the first table.
If it is bigger then 0 then we have passed the first table.
So to fix the bug it intentionally prints the values in reverse and the final product is as expected !
a.k.a backward logic
"""


from docx.api import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm
from docx.shared import Pt
import time


start_time = time.time()

table_count = 0


def add_rows():
    rahshal = Document(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    global table_count
    # Iterate through all tables in the document
    for table in rahshal.tables:
        add_3_rows_and_move_text_down(table)
        make_table_format_as_required(table, rahshal)
        first_row(table)
        style_the_docx_table(table)
        table_count += 1
        print(f"Table {table_count} of {len(rahshal.tables)} modified according to the new format")

    rahshal.save(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    print("All done and save successfully !")
    print(f"--- The code took {time.time() - start_time} seconds to run ---")


def add_3_rows_and_move_text_down(table):
    """
    Adds 3 rows at the bottom of the table and moves all the text down by 3 rows
    """

    # Create a list to store the original values
    original_values = []

    # Iterate through the existing rows and cells in the table starting from row 2
    for i in range(1, len(table.rows)):
        row_values = []
        for cell in table.rows[i].cells:
            row_values.append(cell.text)
        original_values.append(row_values)

    # Insert three new rows
    for _ in range(4):
        table.add_row().cells

    # Iterate through the 'original_values' list and add the content two rows below
    for i, row_values in enumerate(original_values):
        new_row_index = i + 5  # Calculate the new row index
        if new_row_index < len(table.rows):
            for j, cell_value in enumerate(row_values):
                table.cell(new_row_index, j).text = cell_value


def make_table_format_as_required(table, rahshal):
    # Clear the content of the top three rows
    for i in range(4):
        for cell in table.rows[i].cells:
            cell.text = ""

    table.rows[4].cells[0].merge(table.rows[4].cells[1])  # Merge cells 1 and 2
    table.rows[4].cells[2].merge(table.rows[4].cells[3])  # Merge cells 3 and 4
    table.rows[4].cells[4].merge(table.rows[4].cells[5])  # Merge cells 5 and 6

    # Read doc at the top to explanation
    if table_count == 0:
        # Add the default text to row 4
        table.rows[4].cells[0].text = "WGS84 GEO DM"
        table.rows[4].cells[2].text = "WGS84 GEO D"
        table.rows[4].cells[4].text = "ED50 UTM36"
    else:
        table.rows[4].cells[4].text = "WGS84 GEO DM"
        table.rows[4].cells[2].text = "WGS84 GEO D"
        table.rows[4].cells[0].text = "ED50 UTM36"

    set_shading(0, 3, "ffffff", table)  # Set the shading of row 0 to 3 to white
    set_shading(4, 6, "bfbfbf", table)  # Set the shading of row 4 and 5 to light gray

    first_row = table.rows[0]
    table_element = table._tbl
    row_element = first_row._tr
    table_element.remove(row_element)
    merge_third_row(table, rahshal)


def set_shading(row_start: int, row_end: int, color: str, table):
    for row in range(row_start, row_end):  # Set the shading
        for cell in range(len(table.rows[row].cells)):
            # GET CELLS XML ELEMENT
            cell_xml_element = table.rows[row].cells[cell]._tc
            # RETRIEVE THE TABLE CELL PROPERTIES
            table_cell_properties = cell_xml_element.get_or_add_tcPr()
            # CREATE SHADING OBJECT
            shade_obj = OxmlElement("w:shd")
            # SET THE SHADING OBJECT
            shade_obj.set(qn("w:fill"), color)
            # APPEND THE PROPERTIES TO THE TABLE CELL PROPERTIES
            table_cell_properties.append(shade_obj)


def merge_third_row(table, rahshal):
    for cell in range(len(table.rows[2].cells) - 1):
        cell_a = table.cell(2, cell)
        cell_b = table.cell(2, cell + 1)
        merged_cell = cell_a.merge(cell_b)
        merged_cell.text = "קואורדינטות"
        for paragraph in merged_cell.paragraphs:
            for run in paragraph.runs:
                ''' You can enable different designs '''
                # run.font.bold = True
                # run.underline = True
                run.font.size = Pt(12)
                run.font.italic = True


def style_the_docx_table(table):
    table.autofit = False
    # Define the desired column width in centimeters (3 cm)

    # Iterate through the columns and set their widths
    for column in range(len(table.columns)):
        col = table.columns[column]
        col.width = Cm(3.0)

    for row in table.rows:
        for cell in row.cells:
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run.font.name = "Calibri Light"


def first_row(table):
    # Read doc at the top to explanation
    if table_count == 0:
        table.cell(0, 5).text = "שם אזור"
        table.cell(0, 4).text = "מדיניות הגנה"
        table.cell(0, 3).text = "עדיפות"
        table.cell(0, 2).text = "מדיניות הפעלה"
        table.cell(0, 1).text = "עדכון אחרון"
        table.cell(0, 0).text = "הערות"
    else:
        table.cell(0, 0).text = "שם אזור"
        table.cell(0, 1).text = "מדיניות הגנה"
        table.cell(0, 2).text = "עדיפות"
        table.cell(0, 3).text = "מדיניות הפעלה"
        table.cell(0, 4).text = "עדכון אחרון"
        table.cell(0, 5).text = "הערות"

    for cell in table.rows[0].cells:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.italic = True

    set_shading(0, 1, "bfbfbf", table)  # Set the shading of row 0 to light gray


if __name__ == "__main__":
    add_rows()
