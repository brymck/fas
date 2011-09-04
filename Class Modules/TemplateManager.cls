VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TemplateManagEr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pTemplates As New Collection

' Generic collection methods
Public Function Add(ByVal Name As String, _
                    ByVal URL As String, _
                    ByRef Selectors As Variant, _
                    ByRef Abbreviations As Variant) As ConnectionTemplate
    Dim Template As ConnectionTemplate
    Set Template = Find(Name)
    
    If Template Is Nothing Then
        Set Template = New ConnectionTemplate
            
        pTemplates.Add Template, Name
    End If
        
    With Template
        .Name = Name
        .SelectorsFromRange Selectors, Abbreviations
        .URL = URL
    End With
    
    Set Add = Template
End Function
Public Property Get Count() As Long
    Count = pTemplates.Count
End Property
Public Function Item(ByVal Index As Long) As ConnectionTemplate
    Item = pTemplates.Item(Index)
End Function
Public Sub Remove(ByVal Index As Long)
    pTemplates.Remove Index
End Sub

Public Function Find(ByVal Name As String) As ConnectionTemplate
    Dim Template As ConnectionTemplate
    
    On Error Resume Next
    Set Template = pTemplates.Item(Name)
    On Error GoTo 0
    
    Set Find = Template
End Function
