Attribute VB_Name = "ExcelConnections"
Dim ExcelConnection As ADODB.Connection
Dim ExcelRecordSet As ADODB.Recordset

Public Function GetXLAQL(fpath As String) As String
    Dim connStr As String
    
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & fpath _
                    & ";Extended Properties='Excel 12.0;HDR=NO;IMEX=1'"
                    
        '::Both these connection Strings worked, but they would still read in Headers, no matter what I did
'    connStr = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" _
'        & tempPath & ";Extended Properties='Excel 12.0;HDR=No;IMEX=0'"
'    connStr = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & "C:\Users\mdieckman\Desktop\testWB.xlsx" & ";HDR=NO;"
    
    On Error GoTo connErr
    
    Set ExcelConnection = New ADODB.Connection
    ExcelConnection.Open connStr
    
    On Error GoTo BackupAQL  'Attempt to select from the ML Freq Chart first...
    Set ExcelRecordSet = ExcelConnection.Execute("SELECT * FROM [ML Frequency Chart$B7:B7]")
    If ExcelRecordSet.Fields.Count <> 0 Then
        If IsNull(ExcelRecordSet.Fields(0)) Then
            GoTo BackupAQL
        Else
            On Error GoTo NullValErr
            GetXLAQL = ExcelRecordSet.Fields(0)
        End If
    Else
        GoTo BackupAQL
    End If
    
    ExcelConnection.Close
    Exit Function
    
    
BackupAQL:
    On Error Resume Next
    Set ExcelRecordSet = ExcelConnection.Execute("SELECT * FROM [START HERE$I10:I10]")
    GetXLAQL = ExcelRecordSet.Fields(0)
        
    ExcelConnection.Close
    Exit Function
    
connErr:
    ExcelConnection.Close
    Err.Raise Number:=vbObjectError + 1000, Description:="Couldnt Create a connection to the Workbook"

    
NullValErr:
    ExcelConnection.Close
    Err.Raise Number:=vbObjectError + 1200, Description:="Nothing Set or Garbage value set for AQL in Workbook"


End Function


