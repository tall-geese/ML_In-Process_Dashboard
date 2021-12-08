Attribute VB_Name = "Export"
Sub ExportModules()
    Dim wbProj As VBProject
    
    Dim projName As String
    projName = ThisWorkbook.Name
    
    Dim saveDirectory As String
    saveDirectory = InputBox("Add save location", , config.SAVE_PATH)
        
    For Each comp In Workbooks(projName).VBProject.VBComponents
        Dim extension As String
        Dim subFolder As String
        
        Select Case comp.Type
            Case vbext_ComponentType.vbext_ct_Document
                If comp.Name = "ThisWorkbook" Or comp.Name = "Sheet1" Or comp.Name = "Sheet3" Or comp.Name = "Sheet2" Or comp.Name = "Sheet4" Or comp.Name = "Sheet5" Then
                    extension = ".txt"
'                    subFolder = "Excel Objects"
                    subFolder = "Modules"
                    comp.Export (saveDirectory & "\" & comp.Name & extension)
                End If
            Case vbext_ComponentType.vbext_ct_MSForm
                extension = ".frm"
                subFolder = "Modules"
                comp.Export (saveDirectory & "\" & comp.Name & extension)
            Case vbext_ComponentType.vbext_ct_ClassModule
                extension = ".cls"
                subFolder = "Class Modules"
                comp.Export (saveDirectory & "\" & comp.Name & extension)
            Case vbext_ComponentType.vbext_ct_StdModule
                extension = ".bas"
                subFolder = "Modules"
                comp.Export (saveDirectory & "\" & comp.Name & extension)
        End Select
        
       
NextIteration:
       
    Next comp
   
    
End Sub





