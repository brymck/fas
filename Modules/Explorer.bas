Attribute VB_Name = "Explorer"
Option Explicit

Private Manager As New ImportManager

Public Function ImportHTML(ByVal URL As String, _
                           ByVal UpdateFrequency As Long, _
                           ParamArray Selectors() As Variant) As Variant
    Dim Connection As ImportConnection
    Set Connection = Manager.Find(URL)

    If Connection Is Nothing Then
        Set Connection = New ImportConnection
        Connection.URL = URL
        Manager.Add Connection
    Else
        Connection.Refresh
    End If
    
    With Connection
        .Frequency = UpdateFrequency
        .UpdateAll Selectors
    End With
    
    ImportHTML = Connection.Values
End Function

Public Sub CleanUp()
    Manager.CleanUp
End Sub
