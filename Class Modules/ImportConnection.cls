VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Const SECONDS_PER_DAY As Double = 86400

Private pURL As String
Private pIE As SHDocVw.InternetExplorer
Private pSelectors As New Collection
Private pFreq As Long
Private pValues As Variant
Private pNext As Date

Private Sub Class_Initialize()
    ForceUpdate
End Sub

Public Sub UpdateAll(ByVal Selectors As Variant)
    Dim Index As Long
    Dim SelectorLength As Long
    
    SelectorLength = UBound(Selectors)
    Set pSelectors = New Collection
    
    For Index = 0 To SelectorLength
        pSelectors.Add Selectors(Index)
    Next
    
    ForceUpdate
End Sub

Public Sub Refresh()
    pIE.Refresh
    ForceUpdate
End Sub

Private Sub CreateExplorerInstance(ByVal URL As String)
    Dim ShellWindows As New SHDocVw.ShellWindows
    Dim Instance As SHDocVw.InternetExplorer
    
    On Error Resume Next
    For Each Instance In ShellWindows
        If Instance.Document.URL = URL Then
            Set pIE = Instance
            Exit Sub
        End If
    Next
    
    ' Create a new one if nothing else exists
    Set pIE = New SHDocVw.InternetExplorer
    pIE.Navigate URL
End Sub

Public Property Get URL() As String
    URL = pURL
End Property
Public Property Let URL(ByVal Value As String)
    pURL = Value
    CreateExplorerInstance URL
End Property

Public Property Get IE() As SHDocVw.InternetExplorer
    Set IE = pIE
End Property

Public Property Get Document() As MSHTML.HTMLDocument
    Set Document = pIE.Document
End Property

' Generic collection methods
Public Sub Add(ByVal Selector As String)
    pSelectors.Add Selector
    ForceUpdate
End Sub
Public Property Get Count() As Long
    Count = pSelectors.Count
End Property
Public Function Item(ByVal Index As Long) As String
    Set Item = pSelectors.Item(Index)
End Function
Public Sub Remove(ByVal Index As Long)
    pSelectors.Remove Index
End Sub

Public Property Get Values() As Variant
    Dim Index As Long
    Dim SelectorLength As Long
    
    SelectorLength = pSelectors.Count
    ReDim pValues(1 To SelectorLength)
    
    For Index = 1 To SelectorLength
        Dim Value As Variant
        Value = Document.querySelector(pSelectors(Index)).innerText
        If IsNum(Value) Then
            pValues(Index) = CDbl(Value)
        Else
            pValues(Index) = Value
        End If
    Next
    
    Values = pValues
    pNext = Now() + pFreq / SECONDS_PER_DAY
End Property

Public Property Get Frequency() As Long
    Frequency = pFreq
End Property
Public Property Let Frequency(ByVal Value As Long)
    pFreq = Value
End Property

Public Property Get NextUpdate() As Date
    NextUpdate = pNext
End Property

Private Sub ForceUpdate()
    pNext = Now()
End Sub
