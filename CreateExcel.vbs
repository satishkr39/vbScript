Set myxl = createobject("excel.application")
  'To make Excel visible
  myxl.Application.Visible = true 
  myxl.Workbooks.Add

  'Save the Excel file as qtp.xls
  myxl.ActiveWorkbook.SaveAs  "C:\Users\satissingh\Desktop\Test.xlsx"

  'Make sure that you have created an excel file before exeuting the script.
  'Use the path of excel file in the below code
  'Also make sure that your excel file is in Closed state before exeuting the script.

  myxl.Workbooks.Open "C:\Users\satissingh\Desktop\Test.xlsx" 
  myxl.Application.Visible = true

  'this is the name of  Sheet  in Excel file "qtp.xls"   where data needs to be entered 
  set mysheet = myxl.ActiveWorkbook.Worksheets("Sheet1")

  'Enter values in Sheet1.
  'The format of entering values in Excel is excelSheet.Cells(row,column)=value
  mysheet.cells(1,1).value ="Name"
  mysheet.cells(1,2).value ="Age"
  mysheet.cells(2,1).value ="Ram"
  mysheet.cells(2,2).value ="20"
  mysheet.cells(3,1).value ="Raghu"
  mysheet.cells(3,2).value ="15"
  myxl.ActiveWorkbook.Save
   myxl.Application.Quit
  Set myxl=nothing