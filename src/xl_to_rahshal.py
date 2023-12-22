"""
The code copies the coordinates of a specific area from an excel file,
to a table in Microsoft Word file called 'רכשי לב',or for short 'rahshal'
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
import word_functions
import excel_functions
from termcolor import colored


start_time = time.time()


def main():
    """
    The main function loads an Excel workbook and a Word document, searches for specific areas in both
    files, copies data from the Excel file to the Word document, and saves the modified Word document.
    """
    # Load the excel workbook
    excel_workbook_path = r"C:\Users\Daniel\Desktop\Iron Dome\Coordinates.xlsx"
    excel_workbook = openpyxl.load_workbook(excel_workbook_path)

    # Open the active sheet; nz means נ.צ
    nz_xl = excel_workbook.active

    # Load the docx file
    rahshal = Document(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")

    xl_row: int = 1  # Stores the row number of the last area found in excel

    # Get all the tables in the docx file
    tables = rahshal.tables

    # Counts how many tables copied
    table_count = 0

    while xl_row < nz_xl.max_row:
        # Get the next area-name in the order, to be processed
        # Also returns the row where the area was found, so next time the code
        # would not start from the beginning of the excel
        area_name, xl_row = excel_functions.find_area_in_xl(nz_xl, xl_row)

        if area_name != " ":
            # Find the area in the docx file and return the table index.
            # If the area is not in the docx then return None.
            table_index: int = word_functions.find_area_in_rahshal(
                area_name, rahshal, table_count
            )

            # Check if None is returned, meaning that the area is not in the docx
            if table_index is not None:
                # Get a list of all the tables in the docx file
                docx_table = tables[table_index]

                # Update the number of rows in the docx table to suite the new number of coordinates
                word_functions.update_table_dimensions_in_rahshal(
                    rahshal, nz_xl, xl_row, table_index
                )

                # Copy the new coordinates from the excel file to the corresponding docx table
                word_functions.copy_coordinates_from_xl_to_rahshal(
                    nz_xl, docx_table, xl_row, table_count
                )

                # Style the docx table as required (See documentation)
                word_functions.style_the_docx_table(docx_table)

                # Print to the terminal that the table was successfully copied
                print(colored(f"Successfully copied {area_name[::-1]}", "green"))

                table_count += 1

    # Save the modified docx file
    rahshal.save(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    # Close the excel workbook
    excel_workbook.close()

    # End-of-run message
    print("")  # Just a white line
    print("All done and save successfully !")
    print(f"We've copied {table_count} tables from the excel file to rahshal")
    print(f"--- The code took {time.time() - start_time} seconds to run ---")
    print("")  # Just a white line


"""
'TypeError: cannot unpack non-iterable NoneType object' solved by putting all the return values in to one tuple
If the function didn't do it's work then return this one variable as None
At the receiving end, check if the return value is None, else you can safely unpack the tuple and use it.
"""

if __name__ == "__main__":
    main()
