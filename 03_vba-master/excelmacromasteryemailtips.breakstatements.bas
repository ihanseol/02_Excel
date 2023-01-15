Attribute VB_Name = "breakstatements"
Sub breakcombinestatements()
    Sheet1.Activate
    'Use underscore at point to break line.  Space first, underscore, and carriage return.  Also _
    the space first, underscore, and carriage return is different for vba statements.  Look at _
    comments this is a test and Range("A2").Value.
    'Can't use underscore in the middle of an argument name.
    'this is a test _
    is this good
    Range("A1").Value = "break combine statements"
    Range("A2").Value = "break combine " _
                      & "statements"
                      
    'Place multiple statements on the same line
    Range("A3").Value = "Hello": Range("A3").Font.Color = rgbBlue
    'RM:  bad VBA code statements on the same line
End Sub


