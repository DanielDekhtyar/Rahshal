from src.xl_to_rahshal import xl_table_dimensions
import openpyxl
wb = openpyxl.load_workbook(r"C:\Users\Daniel\Desktop\Iron Dome\Coordinates.xlsx")
nz_xl = wb.active  # nz means נ.צ

def test():
        assert xl_table_dimensions(nz_xl , 3) == 22
        assert xl_table_dimensions(nz_xl , 30) == 47
        assert xl_table_dimensions(nz_xl , 83) == 19
        assert xl_table_dimensions(nz_xl , 107) == 16
        assert xl_table_dimensions(nz_xl , 128) == 44
        assert xl_table_dimensions(nz_xl , 177) == 12
        assert xl_table_dimensions(nz_xl , 194) == 31
        assert xl_table_dimensions(nz_xl , 230) == 42
        assert xl_table_dimensions(nz_xl , 277) == 36
        assert xl_table_dimensions(nz_xl , 318) == 14