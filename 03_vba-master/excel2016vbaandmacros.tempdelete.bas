Attribute VB_Name = "tempdelete"
Sub testdelete()
Attribute testdelete.VB_ProcData.VB_Invoke_Func = " \n14"
'
' testdelete Macro
'

'
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "abc"
    Range("M13").Select
    selection.Copy
    Application.CutCopyMode = False
End Sub
