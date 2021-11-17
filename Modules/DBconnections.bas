Attribute VB_Name = "DBconnections"
Dim E10DatabaseConnection As ADODB.Connection
Dim KioskDatabaseConnection As ADODB.Connection
Public ResultRecordSet As ADODB.Recordset
Dim sqlCommand As ADODB.Command
Dim fso As FileSystemObject
Public Enum Connections
    E10 = 0
    Kiosk = 1
End Enum

'****************************************************
'*************  Connection/Query   ******************
'****************************************************

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


Public Function SQLQuery(queryString As String, conn_enum As Connections, params() As Variant)
    Call InitConnection
    Set ResultRecordSet = New ADODB.Recordset
    Set sqlCommand = New ADODB.Command
    With sqlCommand
        .ActiveConnection = GetConnection(conn_enum)
        .CommandType = adCmdText
        .CommandText = queryString
        
        'Params structure
        'params(0) = "jh.JoNum,'NV1452'"
        If (Not params) = -1 Then GoTo 10  'If we have an empty array of parameters
        
        For i = 0 To UBound(params)
            Dim queryParam As ADODB.Parameter
            Set queryParam = .CreateParameter(Name:=Split(params(i), ",")(0), Type:=adVarChar, Size:=255, Direction:=adParamInput, Value:=Split(params(i), ",")(1))
            .Parameters.Append queryParam
        Next i
    End With
    
10
        'TODO Error check here potentially for EOF
    sqlCommand.CommandText = queryString
    ResultRecordSet.Open sqlCommand
    
End Function


'****************************************************
'*************  Public Functions   ******************
'****************************************************

Public Function GetShopLoadInfo() As Variant()
    Set fso = New FileSystemObject
    Dim query As String
    Dim params() As Variant
    
    query = fso.OpenTextFile(Config.QUERY_PATH & "JobLoad.sql", ForReading).ReadAll()
    
    'TODO:set the onError
    Call SQLQuery(queryString:=query, conn_enum:=Connections.E10, params:=params)
    
    'TODO: something here to check EOF
    GetShopLoadInfo = ResultRecordSet.GetRows()

End Function

Public Function GetEpicorCustName(projID As String) As String
    Set fso = New FileSystemObject
    Dim query As String
    Dim params() As Variant

    query = fso.OpenTextFile(Config.QUERY_PATH & "EpicorCustomer.sql", ForReading).ReadAll()
    params = Array("pr.ProjectID," & projID)
    
    Call SQLQuery(queryString:=query, conn_enum:=Connections.E10, params:=params)
    
    If Not ResultRecordSet.EOF Then
        GetEpicorCustName = ResultRecordSet.Fields(0)
    End If
End Function

Public Function GetKioskCustName(cusName As String) As String
    Set fso = New FileSystemObject
    Dim query As String
    Dim params() As Variant

    query = fso.OpenTextFile(Config.QUERY_PATH & "KioskCustomer.sql", ForReading).ReadAll()
    params = Array("ct.Abbreviation," & cusName)
    
    Call SQLQuery(queryString:=query, conn_enum:=Connections.Kiosk, params:=params)
    
    If Not ResultRecordSet.EOF Then
        GetKioskCustName = ResultRecordSet.Fields(0)
    End If
End Function



Public Function GetProductionInfo(jobNum As String, opNum As String) As Variant()
    Set fso = New FileSystemObject
    Dim query As String
    Dim params() As Variant
    
    query = Split(fso.OpenTextFile(Config.QUERY_PATH & "ProductionInfo.sql", ForReading).ReadAll(), ";")(0)
    params = Array("ld.JobNum," & jobNum, "ld.OprSeq," & opNum)
    
    'TODO:set the onError
    Call SQLQuery(queryString:=query, conn_enum:=Connections.E10, params:=params)
    
    'TODO: something here to check EOF
    GetProductionInfo = ResultRecordSet.GetRows()

End Function

Public Function GetProductionInfoSUM(jobNum As String, opNum As String) As Variant()
    Set fso = New FileSystemObject
    Dim query As String
    Dim params() As Variant
    
    query = Split(fso.OpenTextFile(Config.QUERY_PATH & "ProductionInfo.sql", ForReading).ReadAll(), ";")(1)
    params = Array("ld.JobNum," & jobNum, "ld.OprSeq," & opNum)
    
    
    'TODO:set the onError
    Call SQLQuery(queryString:=query, conn_enum:=Connections.E10, params:=params)
    
    'TODO: something here to check EOF
    GetProductionInfoSUM = ResultRecordSet.GetRows()

End Function

Private Function PartMLReady(partNum As String, revNum As String) As Variant
    Set fso = New FileSystemObject
    Dim query As String
    Dim params() As Variant
    
    query = fso.OpenTextFile(Config.QUERY_PATH & "PartMLReady.sql", ForReading).ReadAll()
    params = Array("pr.PartNum," & partNum, "pr.RevisionNum," & revNum)

    Call SQLQuery(queryString:=query, conn_enum:=Connections.E10, params:=params)
    
    PartMLReady = ResultRecordSet.GetRows()

End Function



'****************************************************
'*************  Helper Functions   ******************
'****************************************************



Function IsMeasurLinkJob(JobNumber As String, PartNumber As String, PartRev As String, MachineType As String) As Boolean
    If MachineType = "" Then GoTo 10

    Dim ReadyIndexCol As Collection
    Set ReadyIndexCol = New Collection
    Dim machines() As Variant
    
    'On Error GoTo 10
    
    machines = PartMLReady(partNum:=PartNumber, revNum:=PartRev)
    
    If (Not machines) = -1 Then GoTo 10   'No information for this part, but we may still have created an excel IR for it.
    
    Dim index As Integer
    Dim i As Integer
    
    For i = 0 To UBound(machines)
        If machines(i, 0) = True Then
            If index = 0 Then
                ReadyIndexCol.Add ("")
            Else
                ReadyIndexCol.Add (CStr(index + 1))
            End If
            
        End If
        index = index + 1
    Next i
    
    If ReadyIndexCol.Count = 0 Then GoTo 10
    
    Dim MachineRecordSet As ADODB.Recordset
    Dim MachineQuerySelect As String
    MachineQuerySelect = "SELECT "
    Dim MachineQueryJoins As String
    Dim MachineQueryCriteria As String
    MachineQueryCriteria = " WHERE pr.PartNum = ? AND pr.RevisionNum = ?"
    
    
    For ReadyIndex = 1 To ReadyIndexCol.Count
        MachineQuerySelect = MachineQuerySelect & "ud" & ReadyIndexCol(ReadyIndex) & ".CodeDesc,"
        MachineQueryJoins = MachineQueryJoins & " LEFT OUTER JOIN EpicorLive10.dbo.UDCodes ud" & ReadyIndexCol(ReadyIndex) & " ON pr.ProgramRsrc" & ReadyIndexCol(ReadyIndex) & "_c = ud" & ReadyIndexCol(ReadyIndex) & ".CodeID"
        MachineQueryCriteria = MachineQueryCriteria & " AND ud" & ReadyIndexCol(ReadyIndex) & ".CodeTypeID = 'PGRMRSRC'"
    Next ReadyIndex
    
    MachineQuerySelect = left(MachineQuerySelect, Len(MachineQuerySelect) - 1) & " "
    
    MachineQueryFooter = " FROM EpicorLive10.dbo.PartRev pr " _
    
    Dim machineQuery As String
    machineQuery = MachineQuerySelect & MachineQueryFooter & MachineQueryJoins & MachineQueryCriteria
    Dim params() As Variant
    params = Array("pr.PartNum," & PartNumber, "pr.RevisionNum," & PartRev)
    
    SQLQuery queryString:=machineQuery, conn_enum:=Connections.E10, params:=params
    Set MachineRecordSet = ResultRecordSet
    
    For Each Machine In MachineRecordSet.Fields
        If Machine.Value = MachineType Then
            IsMeasurLinkJob = True
        End If
    Next Machine
                        
10

End Function


















