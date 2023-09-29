from docx import Document
from src.xl_to_rahshal import is_area_in_rahshal

rahshal = Document(r"C:\Users\Daniel\Desktop\Iron Dome\רכשי לב.docx")

def test():
  assert is_area_in_rahshal("חיפה", rahshal) == True
  assert is_area_in_rahshal("נהריה", rahshal) == True
  assert is_area_in_rahshal("מעלות תרשיחא", rahshal) == True
  assert is_area_in_rahshal("כנף 1 (בסיס)", rahshal) == True
  assert is_area_in_rahshal("מגדל העמק", rahshal) == True
  assert is_area_in_rahshal("צפת", rahshal) == True
  assert is_area_in_rahshal("קריית אתא", rahshal) == True
  assert is_area_in_rahshal("שלומי", rahshal) == True
  assert is_area_in_rahshal("מצפה הילה", rahshal) == True
  assert is_area_in_rahshal("מהלול", rahshal) == True
  assert is_area_in_rahshal("עפולה", rahshal) == False
  assert is_area_in_rahshal("טבריה", rahshal) == False
  assert is_area_in_rahshal('מורחב חיפה בט"ש', rahshal) == True