' ******************************************************************************
' Created by Collin Brockman
' Set All Dimensions in Sheet to Fractional Base 16
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
    
    'Get Current Sheet and Sheet Name
    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    Dim swSheetName As String
    swSheetName = swSheet.GetName
        
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
    
    Dim vViews As Variant
    Dim swView As SldWorks.View
    Dim swDispDim As SldWorks.DisplayDimension
    Dim IntStatus As Integer
    
    'All Dimensions to "Overriden"
    Dim vSheets As Variant
    vSheets = swDraw.GetSheetNames
    Dim A As Integer
    For A = 0 To UBound(vSheets)
        Dim iSheet As SldWorks.Sheet
        swDraw.ActivateSheet (vSheets(A))
        Set iSheet = swDraw.GetCurrentSheet()
        Set vViews = iSheet.GetViews
        Dim B As Integer
        For B = 0 To UBound(vViews)
            Set swView = vViews(B)
            Set swDispDim = swView.GetFirstDisplayDimension5()
            While Not swDispDim Is Nothing
                If swDispDim.GetType <> 3 Then
                    Dim UnitInfo As Variant
                    UnitInfo = swDispDim.GetUnits
                    IntStatus = swDispDim.SetUnits2(False, UnitInfo(0), UnitInfo(1), UnitInfo(2), UnitInfo(3), UnitInfo(4))
                    If IntStatus <> 0 Then
                        swApp.SendMsgToUser ("ERROR 1003: Failed to Set Units")
                        Exit Sub
                    End If
                End If
            Wend
        Next
    Next
        
    'Inches = 3
    'Fractions = 2 and Decimal is 1
    'swModel.SetUnits 3, 0, 16, 3, True
    
    
'Rebuild document
swModel.ForceRebuild3 (True)
swModel.GraphicsRedraw2
End Sub
