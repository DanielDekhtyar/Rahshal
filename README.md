# Rahshal

## This is a script designed to copy coordinates from an Excel file to a table in Microsoft Word

### *This script optimizes a time-consuming task in the military* 💂‍♂️

> No classified data is included in this repository

> Author : Daniel Dekhtyar  
> Email : denik2707@gmail.com  
> LinkedIn : https://www.linkedin.com/in/daniel-dekhtyar/  
> 
> Most of this program was written parallel to taking the CS50P class
> (https://github.com/DanielDekhtyar/CS50P)

<br></br>

## Changelog:
> ### Latest update : 22/12/2023
> ### Version : 1.0.4
> _Date format DD-MM-YYYY_

### **[v1.0.4] - 22-12-2023**
---
#### 🔥 Enhancements
- Code is broken into multiple files
    - ***xl_to_rahshal.py*** - The main file that calls the other functions.
    - ***word_functions.py*** - Contains functions that mostly affect the docx file and use the python-docx library
    - ***excel_functions.py*** - Contains functions that mostly affect the Excel file and use the openpyxl library
- README.md updated and now looks as a readme should be :)


### **[v1.0.3] - 29-11-2023**
---
#### 🔥 Enhancements
- sys library not imports as it is not needed
- The success or failure messages are now printed in color (green or red respectfully)
- Error messages are printed in bold
- Comments and documentation made more clear and understandable
- Documentation added to every function
- Code reformated with pylint

### **[v1.0.2] - 3-11-2023**
---
#### 🛠️ Fixed
- first_row_of_table and last_row_of_table initialization changed from None to '0'
#### 🔥 Enhancements
- comments added

### **[v1.0.1] - 4-10-2023**
---
#### 🔥 Enhancements
- Minor code readability improvement

### **[v1.0.0] - 2-10-2023**
---
#### 🚀 Added
- First fully working version of the code.
#### 🎨 Changed
- find_area_in_rahshal() reimplemented to fit the new docx format