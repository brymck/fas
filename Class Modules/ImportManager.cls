VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImportManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pConnections As New Collection

' Generic collection methods
Public Sub Add(ByRef Connection As ImportConnection)
    pConnections.Add Connection
End Sub
Public Property Get Count() As Long
    Count = pConnections.Count
End Property
Public Function Item(ByVal Index As Long) As ImportConnection
    Item = pConnections.Item(Index)
End Function
Public Sub Remove(ByVal Index As Long)
    pConnections.Remove Index
End Sub

Public Function Find(ByVal URL As String) As ImportConnection
    Dim Connection As ImportConnection
    
    ' Find a connection with a matching URL, returning nothing if no match is found
    On Error Resume Next
    For Each Connection In pConnections
        If Connection.URL = URL Then
            Find = Connection
            Exit Function
        End If
    Next
End Function

Public Sub CleanUp()
    Dim ShellWindows As New SHDocVw.ShellWindows
    Dim Instance As SHDocVw.InternetExplorer
    Debug.Print ShellWindows.Count
    
    For Each Instance In ShellWindows
        Instance.Quit
    Next
End Sub
