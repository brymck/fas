Attribute VB_Name = "Templates"
Option Explicit

Private Manager As New TemplateManager

Public Function CustomTemplate(ByVal Name As String, _
                               ByVal URL As String, _
                               ByRef Selectors As Range, _
                               ByRef Abbreviations As Range) As String
    CustomTemplate = Manager.Add(Name, URL, Selectors.Value, Abbreviations.Value).Name
End Function

Public Function ImportTemplate(ByVal Name As String, _
                               ByVal Query As String, _
                               ByVal Frequency As Long, _
                               ParamArray Abbreviations() As Variant) As Variant
    Dim Parsed() As Variant
    Dim Index As Long
    Dim Length As Long
    Length = UBound(Abbreviations)
    
    ReDim Parsed(0 To Length)
    For Index = 0 To Length
        Parsed(Index) = Abbreviations(Index)
    Next
    
    ImportTemplate = Manager.Find(Name).CreateConnection(Query, Frequency, Parsed).Values
End Function
