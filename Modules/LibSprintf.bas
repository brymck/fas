Attribute VB_Name = "LibSprintf"
Option Explicit

Private Const REPLACE_COUNT As Integer = 1

' Mimics sprintf for %s
Public Function sprintf(ByVal Str As String, ParamArray Args()) As String
    Dim Arg As Variant
    
    For Each Arg In Args
        Str = Replace(Str, "%s", Arg, , REPLACE_COUNT)
    Next Arg
    
    sprintf = Str
End Function
