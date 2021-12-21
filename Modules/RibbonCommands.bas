Attribute VB_Name = "RibbonCommands"

                                                    '*************************************************************
                                                    '*************************************************************
                                                    '*                  Ribbon Commands
                                                    '*
                                                    '*************************************************************
                                                    '*************************************************************
Dim ribbonUI As IRibbonUI



Public Sub Ribbon_OnLoad(Ribbon As IRibbonUI)
    Set ribbonUI = Ribbon
    ribbonUI.ActivateTab ("mlDash")
End Sub


Public Sub RefreshShopLoad(ByRef control As IRibbonControl)
    Call Worksheets("ShopLoad").InitShopLoad
End Sub


'Public Sub ShowWarnings(ByRef control As IRibbonControl)
'    Call Worksheets("ShopLoad").CheckAllInspLag
'End Sub
