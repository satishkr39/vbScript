Set myxl = createobject("excel.application") 	' to create excel object
  myxl.Application.Visible = true  'To make Excel visible
  myxl.Workbooks.Add   ' to add a workbook 
  myxl.ActiveWorkbook.SaveAs  "C:\Users\satissingh\Desktop\Test.xlsx"    'Save the Excel file
  myxl.Workbooks.Open "C:\Users\satissingh\Desktop\Test.xlsx"  ' Opening the excel file
  myxl.Application.Visible = true ' making it visible 
  set mysheet = myxl.ActiveWorkbook.Worksheets("Sheet1")  ' setting sheet name 

  'Enter values in Sheet1.
  mysheet.cells(1,1).value ="Name"
  mysheet.cells(1,2).value ="Age"
  mysheet.cells(2,1).value ="Ram"
  mysheet.cells(2,2).value ="20"
  mysheet.cells(3,1).value ="Raghu"
  mysheet.cells(3,2).value ="15"
  myxl.ActiveWorkbook.Save ' saving the changes
   myxl.Application.Quit  ' quit the applicaiton 
  Set myxl=nothing  ' releasing memeory