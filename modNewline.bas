Attribute VB_Name = "modNewline"
' Author: msAshlee
' Date: 2017
' Version: 1
' Written in: Excel 2010 32b VBA
'
'
Public Function nl(ByRef str As String, ParamArray strs() As Variant)
    Dim s As Variant
    For Each s In strs()
       str = str & CStr(s)
    Next s
    str = str & vbCrLf
    nl = str
End Function
