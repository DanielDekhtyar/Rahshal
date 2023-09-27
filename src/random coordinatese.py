import openpyxl as xl
from random import randint

# Create a random coordinate sheet i excel
# It just helped me do the creation of the random coordinates faster
# Don't have a role in the actual program

def main():
  wb = xl.load_workbook(r'C:\Users\Daniel\Desktop\Tests\Coordinates.xlsx')
  sheet = wb.active

  new_row = 0
  for _ in range(10):
    # Create random area length between 4 and 50 and return the last row to last_row
    new_row = create_random_area(sheet, new_row)
    new_row  = new_row + 5 # Add 5 to last_row to get the next area
    
  wb.save(r'C:\Users\Daniel\Desktop\Tests\Coordinates.xlsx')
  print("All done and saved!")


def create_random_area(sheet, start_raw):
  for row in range(start_raw + 4, start_raw + randint(4, 50)):
    for column in range(3, 9):
      cell = sheet.cell(row, column)
      cell.value = randint(1000000, 9999999)   
  return row


main()