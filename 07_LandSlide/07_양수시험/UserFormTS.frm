VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormTS 
   Caption         =   "Time Setting Panel "
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9075.001
   OleObjectBlob   =   "UserFormTS.frx":0000
   StartUpPosition =   1  '소유자 가운데
End
Attribute VB_Name = "UserFormTS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
    Unload Me
End Sub

Function GetNowLast(inputDate As Date) As Date

    Dim dYear, dMonth, getDate As Date

    dYear = Year(inputDate)
    dMonth = Month(inputDate)

    getDate = DateSerial(dYear, dMonth + 1, 0)

    GetNowLast = getDate

End Function

Private Sub ComboBoxFix(ByVal SIGN As Boolean)

    Dim contr As Control
    
    If SIGN Then
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownList
            End If
        Next
    Else
        For Each contr In UserFormTS.Controls
            If TypeName(contr) = "ComboBox" Then
                contr.Style = fmStyleDropDownCombo
            End If
        Next
    End If

End Sub

Private Function whichSection(n As Integer) As Integer

    whichSection = Round((n / 10), 0) * 10

End Function

Private Sub ComboBoxYear_Initialize()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMin As Integer
    
    Dim i, j As Integer
    Dim lastDay As Integer
    
    Dim sheetDate, currDate As Date
    Dim isThisYear As Boolean
    
    sheetDate = Range("c10").Value
    currDate = Now()
    
    If ((Year(currDate) - Year(sheetDate)) = 0) Then
    
        isThisYear = True
        
        nYear = Year(sheetDate)
        nMonth = Month(sheetDate)
        nDay = Day(sheetDate)
        
        nHour = Hour(sheetDate)
        nMin = Minute(sheetDate)
        
    Else
        
        isThisYear = False
        
        nYear = Year(currDate)
        nMonth = Month(currDate)
        nDay = Day(currDate)
        
        nHour = Hour(currDate)
        nMin = Minute(currDate)
            
    End If
    
    
    lastDay = Day(GetNowLast(IIf(isThisYear, sheetDate, currDate)))
    Debug.Print lastDay
    
    For i = nYear - 10 To nYear
        ComboBoxYear.AddItem (i)
    Next i
    
    For i = 1 To 12
        ComboBoxMonth.AddItem (i)
    Next i
    
    For i = 1 To lastDay
        ComboBoxDay.AddItem (i)
    Next i
    
            
    For i = 1 To 12
        ComboBoxHour.AddItem (i)
    Next i
    
    
    
    For i = 0 To 60 Step 10
        ComboBoxMinute.AddItem (i)
    Next i
    
    
    
    ComboBoxYear.Value = nYear
    ComboBoxMonth.Value = nMonth
    ComboBoxDay.Value = nDay
    
    ComboBoxHour.Value = IIf(nHour > 12, nHour - 12, nHour)
    ComboBoxMinute.Value = whichSection(IIf(isThisYear, Minute(sheetDate), Minute(currDate)))
    
   
    If nHour > 12 Then
        OptionButtonPM.Value = True
    Else
        OptionButtonAM.Value = True
    End If
    
    Debug.Print nYear

End Sub

Sub ComboboxDay_ChangeItem(nYear As Integer, nMonth As Integer)
    Dim lasday, i As Integer
    
    lasday = Day(GetNowLast(DateSerial(nYear, nMonth, 1)))
    ComboBoxDay.Clear
    
    For i = 1 To lasday
        ComboBoxDay.AddItem (i)
    Next i
    
    ComboBoxDay.Value = 1

End Sub

Private Sub ComboBoxHour_Change()
    ComboBoxMinute.Value = 0
End Sub

Private Sub ComboBoxMonth_Change()
    '2019-11-26 change
    On Error GoTo Errcheck
    Call ComboboxDay_ChangeItem(ComboBoxYear.Value, ComboBoxMonth.Value)
Errcheck:
        
End Sub

Private Sub EnterButton_Click()
    Dim nYear, nMonth, nDay As Integer
    Dim nHour, nMinute As Integer
    
    Dim nDate, ntime As Date
    
    
    On Error GoTo Errcheck
    nYear = ComboBoxYear.Value
    nMonth = ComboBoxMonth.Value
    nDay = ComboBoxDay.Value
        
    nHour = ComboBoxHour.Value
    nMinute = ComboBoxMinute.Value
            
            
    nHour = nHour + IIf(OptionButtonPM.Value, 12, 0)
            
    nDate = DateSerial(nYear, nMonth, nDay)
    ntime = TimeSerial(nHour, nMinute, 0)
            
    nDate = nDate + ntime
         
    shLongTermTest.Range("c10").Value = nDate
         
Errcheck:
     
    Unload Me
     
End Sub

Private Sub UserForm_Initialize()
    Call ComboBoxYear_Initialize
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub

