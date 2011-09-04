VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ConnectionTemplate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pName As String
Private pSelectors As New Collection
Private pURL As String
Private pUseAbbreviations As Boolean

Public Function CreateConnection(ByVal Query As String, _
                                 ByVal Frequency As Long, _
                                 ByRef Abbreviations() As Variant) As ImportConnection
    Dim Connection As New ImportConnection
    Dim Index As Long
    Dim ParamLength As Long
    
    ParamLength = UBound(Abbreviations)
    
    With Connection
        For Index = 0 To ParamLength
            Dim Abbreviation As Variant
            Abbreviation = Abbreviations(Index)
            .Add pSelectors.Item(Abbreviation)
        Next
        
        .Frequency = Frequency
        .URL = sprintf(pURL, Query)
    End With
    
    Set CreateConnection = Connection
End Function

Public Sub SelectorsFromRange(ByRef Selectors As Variant, _
                              ByRef Abbreviations As Variant)
    Dim Index As Long
    Dim SelectorsLength As Long
    
    SelectorsLength = UBound(Selectors)
    Set pSelectors = New Collection
    
    ' If we're using abbreviations, the params should have the short name
    ' followed by the selector text
    For Index = 1 To SelectorsLength
        pSelectors.Add Selectors(Index, 1), Abbreviations(Index, 1)
    Next
End Sub

Public Property Get Abbreviations() As String()
    Abbreviations = pAbbrevs
End Property

Public Property Get Name() As String
    Name = pName
End Property
Public Property Let Name(ByVal Value As String)
    pName = Value
End Property
Public Property Get UseAbbreviations() As Boolean
    UseAbbreviations = pUseAbbreviations
End Property
Public Property Get URL() As String
    URL = pURL
End Property
Public Property Let URL(ByVal Value As String)
    pURL = Value
End Property
