' ******************************************************************************
' Created by Collin Brockman 7/16/2024
' Set Dimensions Between Leaders to Centered
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
    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet
    
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
    
    'Get all views
    Dim vViews As Variant
    vViews = swSheet.GetViews
    
    'For each veiw
    Dim i As Integer
    For i = 0 To UBound(vViews)
        Dim swView As SldWorks.View
        Set swView = vViews(i)
            'Get first dimension
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = swView.GetFirstDisplayDimension5()
            Dim swDispDimA As Object
            'Iterate through each dimension
            While Not swDispDim Is Nothing
            swDispDim.CenterText = True
                Set swDispDim = swDispDim.GetNext5
            Wend
Next
'Rebuild document
swModel.ForceRebuild3 (True)
End Sub