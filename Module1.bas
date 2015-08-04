Attribute VB_Name = "Module1"
Option Explicit
Public login1 As Integer
Public STR1 As String
Public ct(10)
Public amount As Double
Public prodval(10)
Public fr As String
Public ce(10, 0 To 50) As String
Public cnn As ADODB.Connection
Public i As Integer
Public Sub GetConnected()
Set cnn = New ADODB.Connection
cnn.CursorLocation = adUseClient
cnn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Flipkart.mdb;Persist Security Info=False"
cnn.Open
End Sub
