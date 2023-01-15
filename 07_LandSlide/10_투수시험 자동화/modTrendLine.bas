Attribute VB_Name = "modTrendLine"
Option Explicit
Option Base 0

'Function TLcoef(...) returns Trendline coefficients
'Function TLeval(x, ...) evaluates the current trendline at a given x
'
'The arguments of TLcoef, and the last 4 of TLeval: _
vSheet is the name/number of the sheet containing the chart. _
Use of the name (as in the Sheet        's tab) is recommended _
vCht is the name/number of the chart. To see this, deselect _
the chart, then shift-click it; its name will appear in the _
drop-down list at the left of formula bar. In the case of a _
chart in its own chartsheet, specify this as zero or the zero _
length string "" _
VSeries is a series name/number, and vTL is the series        ' trendline _
number. If the series has a name, it is probably better to _
specify the name. To determine the name/number, as well as _
the trendline number needed for vTL, pass the mouse arrow _
over the trendline. Of course, if there is only one series in _
the chart, you can set vSeries = 1, but beware if you add _
more series to the chart.

'First draft written 2003 March 1 by D J Braden _
Revisions by Tushar Mehta (www.tushar-mehta.com) 2005 Jun 19: _
Various documentation changes _
vCht is now        'optional' _
Correctly handles cases where a term is missing -- e.g., _
y = 2x3 + 3x + 10 _
Correctly handles cases where a coefficient is not shown because _
it is the default value -- e.g., y = Ln(x)+10 _
When only the constant term is present, the original function _
returned it in the correct array element only for the _
polynomial and linear fits. Now, the function returns it in _
the correct array element for other types also. For example, _
for an exponential fit, y=10 will be returned as (10,0) _
Arrays are now base zero.
'Limitations: _
The coefficients are returned to precision *displayed* _
To get the most accurate values, format the trendline label _
to scientific notation with 14 decimal places. (Right-click _
the label to do this) _
Given how XL calculation engine works -- recalculates the _
worksheet first, then the chart(s) -- it is eminently _
possible for the chart to show one trendline and the _
function to return coefficients corresponding to the values _
    shown by the chart *prior* to the recalculation. To see the _
    effect of this        '1 recalculation cycle lag' plot a series of _
    random numbers. _
    An alternative to the functions in this module is the LINEST _
    worksheet function. Except for those few cases where LINEST _
    returns incorrect results, it is the more robust function _
    since it doesn        't suffer from the '1 recalculation cycle' _
    lag. With XL2003 LINEST may even return more accurate _
    results than the trendline.
    
Function TLcoef(vSheet, vCht, vSeries, vTL)
    'To get the coefficients of a chart on a chartsheet, specify vCht _
    as zero or the zero length string ""
    
    'Return coefficients of an Excel chart trendline.  Limitations: See the documentation at the top of the module
    'Note: For a polynomial fit, it is possible the trendline doesn't _
    report all the terms. So this function returns an array of _
    length (1 + the order of the requested fit), *not* the number of _
    values displayed. The last value in the returned array is the _
    constant term; preceeding values correspond to higher-order x.
    
    Dim o As Trendline
    Application.Volatile
    
    If ParamErr(TLcoef, vSheet, vCht, vSeries, vTL) Then Exit Function
    On Error Resume Next
    
    If vCht = "" Or vCht = 0 Then
        If TypeOf Sheets(vSheet) Is Chart Then
            Set o = Sheets(vSheet).SeriesCollection(vSeries).Trendlines(vTL)
        Else
            TLcoef = "#Err: vCht can be omitted only If vSheet Is a " & "chartsheet"
            Exit Function        '*****
        End If
    Else
        Set o = Sheets(vSheet).ChartObjects(vCht).Chart.SeriesCollection(vSeries).Trendlines(vTL)
    End If
    
    On Error GoTo 0
    If o Is Nothing Then
        TLcoef = "#Err: No trendline matches the specified parameters"
    Else
        TLcoef = ExtractCoef(o)
    End If
End Function

Function TLeval(vX, vSheet, vCht, vSeries, vTL)
    'DJ Braden
    'Exp/logs are done for cases xlPower and xlExponential to allow for greater range of arguments.
    
    Dim o As Trendline, vRet
    
    Application.Volatile
    If ParamErr(TLeval, vSheet, vCht, vSeries, vTL) Then Exit Function
    On Error Resume Next

    If vCht = "" Or vCht = 0 Then
        If TypeOf Sheets(vSheet) Is Chart Then
            Set o = Sheets(vSheet).SeriesCollection(vSeries).Trendlines(vTL)
        Else
            TLeval = "#Err: vCht can be omitted only If vSheet Is a " & "chartsheet"
            Exit Function        '*****
        End If
    Else
        Set o = Sheets(vSheet).ChartObjects(vCht).Chart.SeriesCollection(vSeries).Trendlines(vTL)
    End If
    
    On Error GoTo 0
    If o Is Nothing Then
        TLeval = "#Err: No trendline matches the specified parameters"
        Exit Function
    End If

    vRet = ExtractCoef(o)
    If TypeName(vRet) = "String" Then TLeval = vRet: Exit Function
    
    Select Case o.Type
        Case xlLinear
            TLeval = vX * vRet(LBound(vRet)) + vRet(UBound(vRet))
        Case xlExponential        'see comment above
            TLeval = Exp(Log(vRet(LBound(vRet))) + vX * vRet(UBound(vRet)))
        Case xlLogarithmic
            TLeval = vRet(LBound(vRet)) * Log(vX) + vRet(UBound(vRet))
        Case xlPower        'see comment above
            TLeval = Exp(Log(vRet(LBound(vRet))) + Log(vX) * vRet(UBound(vRet)))
        Case xlPolynomial
            Dim Idx As Long
            TLeval = vRet(LBound(vRet)) * vX + vRet(LBound(vRet) + 1)
            For Idx = LBound(vRet) + 2 To UBound(vRet)
                TLeval = vX * TLeval + vRet(Idx)
            Next Idx
    End Select
    
End Function

Private Function DecodeOneTerm(ByVal TLText As String, ByVal SearchToken As String, ByVal UnspecifiedConstant As Byte)
    'splits {optional number}{SearchToken} {optional numeric constant}
    Dim v(1) As Double, TokenLoc As Long
    TokenLoc = InStr(1, TLText, SearchToken, vbTextCompare)
    If TokenLoc = 0 Then
        v(1) = CDbl(TLText)
    Else
        If TokenLoc = 1 Then v(0) = 1 Else: v(0) = Left(TLText, TokenLoc - 1)
        If TokenLoc + Len(SearchToken) > Len(TLText) Then
            v(1) = UnspecifiedConstant
        Else
            v(1) = Mid(TLText, TokenLoc + Len(SearchToken))
        End If
    End If
    DecodeOneTerm = v
End Function


Private Function getXPower(ByVal TLText As String, _
        ByVal XPos As Long)
    If XPos = Len(TLText) Then
        getXPower = 1
    ElseIf IsNumeric(Mid(TLText, XPos + 1, 1)) Then
        getXPower = Mid(TLText, XPos + 1, 1)
    Else
        getXPower = 1
    End If
End Function

Private Function ExtractCoef(o As Trendline)
    Dim XPos        As Long, s As String
    On Error Resume Next
    s = o.DataLabel.Text
    
    On Error GoTo 0
    If s = "" Then
        ExtractCoef = "#Err: No trendline equation found"
        Exit Function        '*****
    End If
    
    If o.DisplayRSquared Then s = Left$(s, InStr(s, "R") - 2)
    s = Trim(Mid(s, InStr(1, s, "=", vbTextCompare) + 1))
    
    Select Case o.Type
        Case xlMovingAvg
        Case xlLogarithmic
            ExtractCoef = DecodeOneTerm(s, "Ln(x)", 0)
        Case xlLinear
            ExtractCoef = DecodeOneTerm(s, "x", 0)
        Case xlExponential
            s = Application.WorksheetFunction.Substitute(s, "x", "")
            ExtractCoef = DecodeOneTerm(s, "e", 1)
        Case xlPower
            ExtractCoef = DecodeOneTerm(s, "x", 1)
        Case xlPolynomial
            Dim lOrd As Long
            ReDim v(o.Order) As Double
            
            s = Application.WorksheetFunction.Substitute(s, " ", "")
            s = Application.WorksheetFunction.Substitute(s, "+x", "+1x")
            s = Application.WorksheetFunction.Substitute(s, "-x", "-1x")
            
            Do While s <> ""
                XPos = InStr(1, s, "x")
                If XPos = 0 Then
                    v(o.Order) = s        'constant term
                    s = ""
                Else
                    lOrd = getXPower(s, XPos)
                    If XPos = 1 Then v(UBound(v) - lOrd) = 1 Else: v(UBound(v) - lOrd) = Left(s, XPos - 1)
                    If XPos = Len(s) Then
                        s = ""
                    ElseIf IsNumeric(Mid(s, XPos + 1, 1)) Then
                        s = Trim(Mid(s, XPos + 2))
                    Else
                        s = Trim(Mid(s, XPos + 1))
                    End If
                End If
            Loop
            
            ExtractCoef = v
    End Select
    
End Function


Private Function ParamErr(v, ParamArray parms())
    Dim l           As Long
    For l = LBound(parms) To UBound(parms)
        If VarType(parms(l)) = vbError Then
            v = parms(l)
            ParamErr = True
            Exit Function
        End If
    Next l
End Function

