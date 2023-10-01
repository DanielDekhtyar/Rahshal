import openpyxl
from src.xl_to_rahshal import find_area_in_xl
excel_workbook = openpyxl.load_workbook(r"C:\Users\Daniel\Desktop\Iron Dome\Coordinates.xlsx")
nz_xl = excel_workbook.active  # Open the active sheet; nz means נ.צ

def test():
  assert find_area_in_xl(nz_xl, 1) == ("חיפה", 3)
  assert find_area_in_xl(nz_xl, 3) == ("נהריה", 34)
  assert find_area_in_xl(nz_xl, 34) == ("מעלות תרשיחא", 86)
  assert find_area_in_xl(nz_xl, 86) == ("כנף 1 (בסיס)", 110)
  assert find_area_in_xl(nz_xl, 110) == ("צפת", 131)
  assert find_area_in_xl(nz_xl, 131) == ("מגדל העמק", 148)
  assert find_area_in_xl(nz_xl, 148) == ("עפולה", 197)
  assert find_area_in_xl(nz_xl, 197) == ("קריית אתא", 214)
  assert find_area_in_xl(nz_xl, 214) == ("שלומי", 250)
  assert find_area_in_xl(nz_xl, 250) == ("מהלול", 297)
  assert find_area_in_xl(nz_xl, 297) == ("מצפה הילה", 338)
  assert find_area_in_xl(nz_xl, 338) == ("טבריה", 356)