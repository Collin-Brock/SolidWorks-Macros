' ******************************************************************************
' Created by Collin Brockman
' Sets Sheet Format
' ******************************************************************************
'Preconditions:
'Open Drawing
Option Explicit
Dim swApp As SldWorks.SldWorks
Sub main()
    'Setup current drawing and sheet
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Set swModel = swApp.ActiveDoc
    Set swDraw = swModel
    
    'Check if the instance is running
    If swModel Is Nothing Then
        swApp.SendMsgToUser ("ERROR 1002: No Open File")
        Exit Sub
    End If
    
    'Check if the File type is a Drawing
    If swModel.GetType <> swDocumentTypes_e.swDocDRAWING Then
        swApp.SendMsgToUser ("ERROR 1001: Wrong File Type")
        Exit Sub
    End If
    
    'Get All Sheets by Name
    Dim swSheetNames As Variant
    swSheetNames = swDraw.GetSheetNames
    
    'For Loop Over each Sheet
    Dim i As Integer
    For i = 0 To UBound(swSheetNames)
        Dim BoolStatus As Boolean
        BoolStatus = swDraw.ActivateSheet(swSheetNames(i))
        If BoolStatus = True Then
            swApp.SendMsgToUser ("Error 1003: Sheet Not Activated")
            Exit Sub
        End If
        Dim swSheet As SldWorks.Sheet
        Set swSheet = swDraw.GetCurrentSheet
        
        'Get Sheet Properties and set sheet properties to the same except the template is updated *Notes are Deleted*
        Dim swSheetProp As Variant
        swSheetProp = swSheet.GetProperties2
        BoolStatus = swDraw.SetupSheet5(swSheetNames(i), swSheetProp(0), 12, swSheetProp(2), swSheetProp(3), swSheetProp(4), Path , swSheetProp(5), swSheetProp(6), swSheetProp(7), True)
        swSheet.ReloadTemplate (False)
    Next
    
'Rebuild document
swModel.ForceRebuild3 (True)
End Sub