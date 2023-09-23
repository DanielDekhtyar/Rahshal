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
            print(area_name)
        
    print(f"--- The code took {time.time() - start_time} seconds to run ---")


def find_area_in_xl(nz_xl, xl_row): # TESTED AND DONE !
    for row in range(xl_row + 1, nz_xl.max_row + 1):
        cell_value = nz_xl[f"A{row}"].value
        if cell_value is not None:
            if "קורדינטות" not in cell_value:
                return cell_value, row
    return " ", row

if __name__ == "__main__":
    main()