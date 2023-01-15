Attribute VB_Name = "Chapter09"
Sub openworkbook()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("9").Activate
    Workbooks.Open ("G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\namesoriginal.xlsm")
    Workbooks.OpenText Filename:="G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\Grades1988-1990.txt", DataType:=xlDelimited, Tab:=True
    'ThisWorkbook property used to save Excel with macros
    ThisWorkbook.SaveAs Filename:="G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\excelprogramming.xlsm"
    'Save .xlsx Excel file
    'ActitveWorkbook.Save Filename:="G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\excelprogramming.xlsm"
End Sub

Sub closeworkbook()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("9").Activate
    Workbooks.Open ("G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\namesoriginal.xlsm")
    'Workbooks("namesoriginal.xlsm").Close 'doesn't work for Excel with Macro
    Workbooks("namesoriginal.xlsm").Close Savechanges:=True, Filename:="namesoriginal.xlsm", Routeworkbook:=False
End Sub

Sub newworkbook()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("9").Activate
    'create blank new Excel Workbook
    Workbooks.Add (xlWBATWorksheet)

    'Book1.xlsx file must exist.  Saves Book1.xlsx file as savesfilename.xlsx
    Dim newworkbook As Workbook
    Set newworkbook = Workbooks.Add("G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\Book1.xlsx")
    newworkbook.title = "workbook Title"
    newworkbook.SaveAs "saveasfilename.xlsx"
End Sub

Sub deleteworkbook()
    Application.Workbooks("excelprogramming.xlsm").Worksheets("9").Activate
    'Book1.xlsx file must exist.  Saves Book1.xlsx file as savesfilename.xlsx
    Dim newworkbook As Workbook
    Set newworkbook = Workbooks.Add("G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\Book1.xlsx")
    newworkbook.title = "workbook Title"
    newworkbook.SaveAs "G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\filetobedeleted.xlsx"
    Workbooks("filetobedeleted.xlsx").Close
    'delete excel file
    Kill "G:\Raymond\Excel Files 2GB Backup 040118\VBA Macros Round Two\filetobedeleted.xlsx"
    Range("A1").Value = CurDir
End Sub

