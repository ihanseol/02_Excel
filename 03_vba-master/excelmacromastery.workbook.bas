Attribute VB_Name = "workbook"
Sub basicworkbook()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("workbook").Activate
    'RM:  I looked at most of the basics
    'Workbooks("excelmacromastery.xlsm") access open workbook filename
    'Workbooks(1) access first open workbook
    'workbooks(workbooks.Count) access last open workbook
    'ActiveWorkbook access active workbook
    'ThisWorkbook access workbook containing vba code
    'dim declareworkbookvariable as workbook
    'set declareworkbookvariable = Workbooks("filename.xlsx")
    'set declareworkbookvariable = ThisWorkbook
    'set declareworkbookvariable = Workbooks(1)
    'declareworkbookvariable.Activate activate workbook
    'set declareworkbookvariable = Workbooks.Add create new workbook
    'set declareworkbookvariable = Workbooks.Open("c:\docs\filename.xlsx") 'open workbook
    'declareworkbookvariable.Save save workbook
    'declareworkbookvariable.SaveCopyAs "C:\copy.xlsx" 'save workbook copy
    'FileCopy "C:\file1.xlsx","C:\Copy.xlsx" copy workbook if closed
    'declareworkbookvariable.SaveAs "Backup.xlsx" saveas workbook
    'Workbooks("filename.xlsx").Close closes workbook
End Sub

Sub writetoworkbook()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("workbook").Activate
    Workbooks("excelmacromastery.xlsm").Worksheets("workbook").Range("A1") = "Write A1 here"
End Sub

Sub workbookproperties()
    Application.Workbooks("excelmacromastery.xlsm").Worksheets("workbook").Activate
    Range("A2").Value = Workbooks.count          'prints number of open workbooks
    Range("A3") = Workbooks("excelmacromastery.xlsm").name 'prints file name excelmacromastery.xlsm
    Range("A4") = Workbooks("excelmacromastery.xlsm").Worksheets.count 'prints number of worksheets in excelmacromastery.xlsm
    Range("A5") = Workbooks("excelmacromastery.xlsm").ActiveSheet.name 'prints active sheet workbook in excelmacromastery.xlsm
    'Workbooks("excelmacromastery.xlsm").close 'closes excelmacromastery.xlsm
End Sub

