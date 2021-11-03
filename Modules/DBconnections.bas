Attribute VB_Name = "DBconnections"
Dim E10DatabaseConnection As ADODB.Connection
Dim KioskDatabaseConnection As ADODB.Connection
Public ResultRecordSet As ADODB.Recordset
Dim sqlCommand As ADODB.Command
Public Enum Connections
    E10 = 0
    Kiosk = 1
End Enum


Private Sub Workbook_BeforeClose(Cancel As Boolean)
    'Close connection before closing workbook
    On Error Resume Next
    JobRecordSet.Close
    E10DatabaseConnection.Close
End Sub


Private Sub InitConnection()
    'Initialize E10 Connection on startup
    If E10DatabaseConnection Is Nothing Then
    
        Set E10DatabaseConnection = New ADODB.Connection
        E10DatabaseConnection.ConnectionString = Config.E10_CONN_STRING
        E10DatabaseConnection.Open
        
    End If
    If KioskDatabaseConnection Is Nothing Then
    
        Set KioskDatabaseConnection = New ADODB.Connection
        KioskDatabaseConnection.ConnectionString = Config.KIOSK_CONN_STRING
        KioskDatabaseConnection.Open
        
    End If

End Sub

Private Function GetConnection(conn_enum As Connections) As ADODB.Connection
    Select Case conn_enum
        Case 0
            Set GetConnection = E10DatabaseConnection
        Case 1
            Set GetConnection = KioskDatabaseConnection
        Case Else
    End Select
End Function


Public Function SQLQuery(queryString As String, conn_enum As Connections)
    Call InitConnection
    Set ResultRecordSet = New ADODB.Recordset
    Set sqlCommand = New ADODB.Command
    sqlCommand.ActiveConnection = GetConnection(conn_enum)
    
    sqlCommand.CommandText = queryString
    ResultRecordSet.Open sqlCommand

End Function
