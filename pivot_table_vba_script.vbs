Dim xl1
Set xl1 = CreateObject("Excel.Application")
xl1.DisplayAlerts = False
xl1.Application.run "'C:\Users\holde\Documents\Python\code\XML - XLSX - JSON\pivotmacro.xlsm'!pivot"
xl1.Application.Quit
Set xl1 = Nothing  