Attribute VB_Name = "Chapter06"
Sub namesinvba()
    Application.Workbooks("excel2016vbaandmacros.xlsm").Worksheets("6").Activate
    Names("NewnameFruitslocal").Delete        'Delete a name
    ActiveWorkbook.Names.Add Name:="Fruits", RefersTo:="6!A1:F6"        'This creates a global name Fruits
    'Worksheets("6").Names.Add Name:="Fruitslocal", RefersTo:="6!A1:F6" 'This creates a local name Fruitslocal
    'same as
    Range("A1:F6").Name = "Fruitslocal"        'Actually names cells Fruitslocal
    Names("Fruitslocal").Name = "NewnameFruitslocal"
    Names("NewnameFruitslocal").Comment = "Comment appear in Name Manager"
    'The most common use of names is for storing ranges; however, names can also assign names
    'to name formulas, strings, numbers, and arrays, as described in the following pages.
    'RM:  skipped section.
End Sub
