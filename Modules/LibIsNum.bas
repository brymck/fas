Attribute VB_Name = "LibIsNum"
Option Explicit

Public Function IsNum(ByVal Value As String, _
                      Optional ByVal RespectLocale As Boolean = False) As Boolean
    Dim DecimalPoint As String
    Dim ThousandsSeparator As String
  
    If RespectLocale Then
        DecimalPoint = Format$(0, ".")
        ThousandsSeparator = Mid$(Format$(1000, "#,###"), 2, 1)
    Else
        DecimalPoint = "."
        ThousandsSeparator = ","
    End If
  
    ThousandsSeparator = Mid$(Format$(1000, "#,###"), 2, 1)
    Value = Replace$(Value, ThousandsSeparator, "")
  
    If Value Like "[+-]*" Then
        Value = Mid$(Value, 2)
    End If
  
    IsNum = Not Value Like "*[!0-9" & DecimalPoint & "]*" And _
            Not Value Like "*" & DecimalPoint & "*" & DecimalPoint & "*" And _
            Len(Value) > 0 And _
            Value <> DecimalPoint
End Function


