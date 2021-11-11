VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ProdHrsChart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public shape As Excel.shape
Public WithEvents cht As Excel.Chart
Attribute cht.VB_VarHelpID = -1

'Private Sub cht_Select(ByVal ElementID As Long, ByVal Arg1 As Long, ByVal Arg2 As Long)
'    MsgBox "This was selected"
'End Sub
'
'Private Sub cht_Activate()
'    MsgBox "this was activated"
'
'
'End Sub

'Private Sub cht_BeforeDoubleClick(ElementID As Long, Arg1 As Long, Arg2 As Long, Cancel As Boolean)
'    MsgBox "this was double clicked"
'End Sub


'Private Sub cht_BeforeDoubleClick(ByVal ElementID As Long, _
'                                    ByVal Arg1 As Long, _
'                                    ByVal Arg2 As Long, _
'                                    ByRef Cancel As Boolean)
'    MsgBox "this was double clicked"
'End Sub
