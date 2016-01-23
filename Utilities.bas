Attribute VB_Name = "Utilities"
'
'
Sub ValidateRange(x)
'This sets x to Nothing if x has not yet been set
'The error handling below appears to be the only way to determine this
'We encapsulate error handling here so that other functions
'can simply test whether x is Nothing
Dim Temp As Range
Set Temp = Nothing
On Error Resume Next
Set Temp = Range(x.Address)
On Error GoTo 0
Set x = Temp
End Sub
'
'
Function RangeAddress(x As Object, s As String) As String
'This is called with s="". So it returns "" if x has not been set
'else it returns the address of x
If x Is Nothing Then
    RangeAddress = s
Else
    RangeAddress = x.Address(External:=True)
End If
End Function

'
'
Function myFormat(x As Double, _
         Optional ForceZeroInf As Boolean = False) As String
'Format the number in fixed or exponential format as appropriate
'If the optional variable ForceZeroInf is True then
'very small numbers are displayed as zero and
'very large numbers are displayed as +/- infinity
Const Small As Double = 0.0001, Large As Double = 1000, _
      vSmall As Double = 1E-30, vLarge As Double = 1E+30
Dim AbsX As Double, s As String
AbsX = Abs(x)
'begin with exponential format
fmt = "\ 0.0000E+00;-0.0000E+00"
'but if numbers are within range where
'change to fixed decimal format
If (AbsX > Small And AbsX < Large - Small / 2) Then
    fmt = "\ ###0.0000\ \ \ \ ;-###0.0000\ \ \ \ "
End If
'format the number into string
s = Format(x, fmt)
'If the optional variable ForceZeroInf is True then
'very small numbers are displayed as zero and
If (ForceZeroInf And AbsX <= vSmall) Then
    s = "0      "
End If
'If the optional variable ForceZeroInf is True then
'very large numbers are displayed as +/- infinity
If (ForceZeroInf And AbsX >= vLarge) Then
    If (x < 0) Then
        s = "+Infinity   "
    Else
        s = "-Infinity   "
    End If
End If
'put the formatted number right justified in a field of width 13
myFormat = Space(13)
RSet myFormat = s
End Function
'
'
Function MultiStrings_0(Separator As String, ForceZeroInf As Boolean, _
         ParamArray StrArr() As Variant _
         ) As String
'Given an array of arguments in StrArr,
'concatenate them into a single string
'with a Separator between two arguments
'Numeric arguments are formatted using myFormat
'The parameter ForceZeroInf is passed on to myFormat
'String and numeric arguments can be interspersed
'This should not be called directly, because then the
'elements of StrArr are in StrArr(0), StrArr(1) etc
'When called indirectly the elements of StrArr are in
'StrArr(0)(0), StrArr(0)(1) etc
'User interface into this routine is through MultiStrings,
'MultiLine, SingleLine etc.
MultiString_0 = ""
For i = 0 To UBound(StrArr(0))
    If (TypeName(StrArr(0)(i)) = "Double") Then
        MultiStrings_0 = MultiStrings_0 & _
                      myFormat(val(StrArr(0)(i)), ForceZeroInf)
    Else
        MultiStrings_0 = MultiStrings_0 & StrArr(0)(i)
    End If
    If (i < UBound(StrArr(0))) Then _
        MultiStrings_0 = MultiStrings_0 & Separator
Next i
End Function
'
'
Function MultiStrings(Separator As String, ForceZeroInf As Boolean, _
         ParamArray StrArr() As Variant _
         ) As String
'Interface to MultiStrings_0
'This needed because MultiStrings_0 should not be called directly
MultiStrings = MultiStrings_0(Separator, ForceZeroInf, StrArr)
End Function
'
'
Function MultiLine(ParamArray StrArr() As Variant) As String
'Interface to MultiStrings_0 with 2 new lines as the Separator
'The paramater ForceZeroInf into myFormat is set to false
MultiLine = MultiStrings_0(Chr$(10) & Chr$(10), False, StrArr)
End Function
'
'
Function SingleLine(ParamArray StrArr() As Variant) As String
'Interface to MultiStrings_0 with a space as the Separator
'The paramater ForceZeroInf into myFormat is set to false
SingleLine = MultiStrings_0(" ", False, StrArr)
End Function
'
'
Function MultiLineZ(ParamArray StrArr()) As String
'Interface to MultiStrings_0 with 2 new lines as the Separator
'The paramater ForceZeroInf into myFormat is set to true
MultiLineZ = MultiStrings_0(Chr$(10) & Chr$(10), True, StrArr)
End Function
'
'
Function SingleLineZ(ParamArray StrArr()) As String
'Interface to MultiStrings_0 with a space as the Separator
'The paramater ForceZeroInf into myFormat is set to true
SingleLineZ = MultiStrings_0(" ", True, StrArr)
End Function

Function Reverse_Array(x As Range)
Dim r As Integer, c As Integer, i As Integer
Dim outarray() As Variant
c = x.Columns.Count
r = x.Rows.Count
If (c <> 1 And r <> 1) Then
    MsgBox ("Array to be reversed must be either a row or a column")
    Exit Function
End If
ReDim outarray(r, c)
If r = 1 Then
    For i = 1 To c
        outarray(0, c - i) = x(1, i)
    Next i
Else
    For i = 1 To r
        outarray(r - i, 0) = x(i, 1)
    Next i
End If
Reverse_Array = outarray
End Function


