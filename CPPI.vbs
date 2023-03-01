

Dim xlApp 
Dim xlBook 
Dim fso
Set fso = CreateObject("Scripting.FileSystemObject")
Dim fullpath
fullpath = fso.GetAbsolutePathName(".")
Set xlApp = CreateObject("Excel.Application") 
' Set xlBook = xlApp.Workbooks.Open(fullpath &"\CPPI.xlsm")
xlApp.Application.Run fullpath &"\CPPI.xlsm!CPPI.CPPI_mrkt_scenario"


' xlBook.Save
' xlBook.Close
xlApp.Quit



