## Excel Web Scraper - Address Lookup

### Description
Written entirely in Excel with VBA and XML/HTTP Object Libraries
Performs searches by license numbers and pulls the address out of the webpage DOM 
and pastes it in the spreadsheet

### Installation Instructions
This project was ritten in Microsoft Excel 2016 edition, it has not been tested in older versions.<br/>
It will not work in iOS version of Excel due to missing object libraries

1. Open a blank Excel workbook
2. Open the VBA editor, there are several ways to get to it:
    - Alt + F11
    - Go to File > Options > Customize Ribbon 
        1. Click Developer check box 
        2. Click OK 
        3. Click Developer in your ribbon
        4. All the way on the left click the "Visual Basic" icon
3. In the VBA editor, go to Tools > References and make sure the following libraries are enabled:
  Visual Basic For Applications
  Microsoft Excel 16.0 Object Library
  Microsoft Office 16.0 Object Library
  OLE Automation
  Microsoft HTML Object Library
  Microsoft Internet Controls
  Microsoft XML v6.0
4. Copy all the code in the "Main" file of this repository, and paste it into Module 1
5. Copy the list of license numbers in the "Licenses" file, and paste it in column A beginning in cell A1
6. Run the macro, there are several ways to do it:
    - In the VBA edtior with Module 1 window active you can
        - Press F5 to execute all code, or press F8 continuously to step through the code
        - Press the play button at the top of the VBA editor
    - Back in the spreadsheet in the Developer ribbon 
        1. Click on "Insert" 
        2. Click "Button" (it should be the first one)
        3. Create a button somewhere in the spreadsheet 
        4. Right click on the button and select "Assign Macro"
        5. Select the macro that we pasted into Module 1 and click OK 
        6. Now you can click the button to run the macro
7. Watch the addresses populate the spreadsheet!
