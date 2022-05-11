set obj = CreateObject("Excel.Application")
obj.visible = True
set obj1 = obj.WorkBooks.open("C:\Users\satissingh\Documents\Important\VBScripting\Test.xlsx")
set obj2 = obj1.WorkSheets("Sheet1")
obj2.Rows("4:4").Delete    		' Deleting 4th row in excel     
obj1.save()
obj1.close
obj.Quit
set obj1 = Nothing
set obj2 = Nothing
