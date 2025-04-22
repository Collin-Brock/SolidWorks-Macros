' ******************************************************************************
' Created by Collin Brockman 7/16/2024
' Reattach Selected Dimensions
' ******************************************************************************
'Start Function no Parameters in or out
Option Explicit
Sub main()
    'Initalize App and File to current instance
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
   
    'Check if more than two object is selected (view and component)
    If swSelMgr.GetSelectedObjectCount2(0) > 1 Then
        swApp.SendMsgToUser ("Error 1004: More Object Than One Selected")
        Exit Sub
    End If
    'Check if less than one object is selected (view)
    If swSelMgr.GetSelectedObjectCount2(0) < 1 Then
        swApp.SendMsgToUser ("Error 1005: Nothing Selected")
        Exit Sub
    'Check if object type is view component
    End If
    If swSelMgr.GetSelectedObjectType3(1, 0) <> 14 Then
        swApp.SendMsgToUser ("Error 1006: Incorrect Object Selected")
        Exit Sub
    End If
    'Run the Reattach command on selected dimension
    Dim Status As Variant
    Status = swApp.RunCommand(swCommands_Reattach_Dim, "REATTACH MACRO")
End Sub

