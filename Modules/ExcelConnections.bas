Attribute VB_Name = "ExcelConnections"
Dim ExcelConnection As ADODB.Connection
Dim ExcelRecordSet As ADODB.Recordset

Public Function openXl()
    Dim connStr As String
    connStr = "Driver={Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)};DBQ=" & "C:\Users\mdieckman\Desktop\testWB.xlsx;HDR=NO;IMEX=0"
'    connStr = "Provider=MSDASQL.1;DSN=Excel Files;DBQ=" & "C:\Users\mdieckman\Desktop\testWB.xlsx" & ";HDR=NO;"
    
    
'    connStr = "DRIVER={Microsoft Excel Driver (*.xls)};ReadOnly=1;DBQ=" & "C:\Users\mdieckman\Desktop\testWB.xlsx"
'    connStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\mdieckman\Desktop\testWB.xlsx;Extended Properties=Excel 8.0"
    Set ExcelConnection = New ADODB.Connection
    ExcelConnection.Open connStr
'    Set ExcelRecordSet = ExcelConnection.Execute("SELECT * FROM [Sheet1$A1:B1]")
    Set ExcelRecordSet = ExcelConnection.Execute("SELECT * FROM [Sheet1$A1:A3]")
    
    i = 0
    While Not ExcelRecordSet.EOF
        MsgBox ExcelRecordSet.Fields(i)
        ExcelRecordSet.MoveNext
    Wend
    ExcelConnection.Close
    Exit Function
    
closeWB:
    ExcelConnection.Close

End Function



'
'Dim cn As ADODB.Connection
'  Set cn = New ADODB.Connection
'  With cn
'    .Provider = "MSDASQL"
'    .ConnectionString = "Driver={Microsoft Excel Driver (*.xls)};" & _
'  "DBQ=C:\MyFolder\MyWorkbook.xls; ReadOnly=False;"
'    .Open
'  End With
'strQuery = "SELECT * FROM [Sheet1$A1:B10]"
