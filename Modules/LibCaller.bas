Attribute VB_Name = "LibCaller"
Option Explicit

Public Function SizeToCaller(ByRef Arr As Variant)
    ' Dimension array to size of calling rows and columns
    With Application.Caller
        If .Rows.Count > .Columns.Count Then
            CallerMax = .Rows.Count
            ReDim Arr(1 To CallerMax, 1 To 1)
        Else
            CallerMax = .Columns.Count
            ReDim Arr(1 To 1, 1 To CallerMax)
        End If
    End With
End Function
