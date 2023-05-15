Attribute VB_Name = "ExcelHelpers"
'*************************************************************************************************
'
'   ExcelHelpers
'       For Interacting with other Microsoft Office objects outside of ThisWorkbook
'       1. GetAQL  - from the inspection report workbook.
'*************************************************************************************************


Public Function GetAQL(aqlVal As String, ProdQty As Integer, aqlWB As Workbook) As String
    Dim reqQty As String
    Dim row As String
    Dim col As Integer
    Dim returnAQL As String
 
    If aqlVal = "100%" Then
        GetAQL = CStr(ProdQty)
        Exit Function
    End If
    
    Select Case ProdQty
        Case 2 To 4
            row = "2"
        Case 5 To 10
            row = "3"
        Case 11 To 15
            row = "4"
        Case 16 To 20
            row = "5"
        Case 22 To 25
            row = "6"
        Case 26 To 30
            row = "7"
        Case 31 To 50
            row = "8"
        Case 51 To 90
            row = "9"
        Case 91 To 150
            row = "10"
        Case 151 To 280
            row = "11"
        Case 281 To 500
            row = "12"
        Case 501 To 1200
            row = "13"
        Case 1201 To 3200
            row = "14"
        Case 3201 To 32000
            row = "15"
        Case Else
            GoTo ProdQtyErr
    End Select
    
    With aqlWB.Worksheets("AQL")
        col = Application.WorksheetFunction.Match(CDbl(aqlVal), .Range("A1:J1"), 0)
        reqQty = .Range(GetAddress(col) & row).Value
    End With
    
    'sometimes The qty required by an AQL is greater than the amount of parts we've made for some reason
    'Like for 10 parts with an AQL of 1.00
    If reqQty > ProdQty Then
        returnAQL = CStr(ProdQty)
    Else
        returnAQL = CStr(reqQty)
    End If
    
    GetAQL = returnAQL
    
    Exit Function
    
ProdQtyErr:
    result = MsgBox("There was a problem attempting to interpret this job's production quantity of " & ProdQty & vbCrLf & _
                     "Verify that this qty is correct in Epicor and contact a QE for assistance.", vbExclamation)
    
End Function



Public Function GetAddress(column As Integer) As String
    Dim vArr
    vArr = Split(Cells(1, column).Address(True, False), "$")
    GetAddress = vArr(0)

End Function




