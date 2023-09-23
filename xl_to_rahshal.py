import openpyxl
from docx import Document
import time

start_time = time.time()

def main():
    # Load the excel workbook
    wb = openpyxl.load_workbook(r"C:\Users\Daniel\Desktop\Iron Dome\Coordinates.xlsx")
    nz_xl = wb.active  # nz means נ.צ
    rahshal = Document(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")
    xl_row = 1
    while xl_row < nz_xl.max_row:
        area_name, xl_row = find_area_in_xl(nz_xl, xl_row)
        if area_name != " ":
            results = is_area_in_rahshal(area_name, rahshal)
            if results is not None: # Check if None is returned, meaning that the area is not in the docx
                area_name, paragraph = results
        
    print(f"--- The code took {time.time() - start_time} seconds to run ---")


def find_area_in_xl(nz_xl, xl_row): # TESTED AND DONE !
    for row in range(xl_row + 1, nz_xl.max_row + 1):
        cell_value = nz_xl[f"A{row}"].value
        if cell_value is not None:
            if "קורדינטות" not in cell_value:
                return cell_value, row
    return " ", row


def is_area_in_rahshal(area_name, rahshal):
    for paragraph in enumerate(rahshal.paragraphs):
        if area_name == paragraph[1].text:
            results = area_name, paragraph
            return results
        
    print(f"The area {area_name[::-1]} is not in the docx") # If area is not in rahshal, an error message will appear
    return None
# 'TypeError: cannot unpack non-iterable NoneType object' solved by putting all the return values in to one tuple
# If the function didn't do it's work then return this one variable as None
# At the receiving end, check if the return value is None, else you can safely unpack the tuple and use it.

if __name__ == "__main__":
    main()