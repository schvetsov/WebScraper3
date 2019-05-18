## Excel Web Scraper - Address Lookup

### Description
Written entirely in Excel with VBA and XML/HTTP Object Libraries
Performs searches by license numbers and pulls the address out of the webpage DOM 
and pastes it in the spreadsheet

### Installation Instructions
This web scraper was written in Microsoft Excel 2016 edition, it might work in older versions.
To use, follow these instructions:

1. Open a blank Excel workbook
2. Open the VBA editor: Tools > Macro > Visual Basic Editor
3. Go to Tools > References and enable the following libraries:
  HTML Object Library
  XML Object Library
4. Copy all the code in this repository in the "Main" file, and paste it into Module 1
5. Copy the list of license numbers, and paste it in column A
6. Run the macro by pressing F5 in the VBA editor, pressing the play button in the VBA editor, 
  or assigning the macro to a button in the spreadsheet and pressing it.
7. Watch the addresses populate the spreadsheet
