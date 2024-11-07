' ******************************************************************************
' Created by Collin Brockman
' Set All Dimensions in Sheet to Fractional
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
            'Iterate through each dimension
            While Not swDispDim Is Nothing
                'Check that selected dimension is NOT and angular dimension
                If swDispDim.GetType <> 3 Then
                    'Set each display dimension to Fractional
                    Dim IntStatus As Integer
                    'Call is (False , Unit Display , FBase , Denominator of Fraction , Round T/F , DRound (only if FBase is 1))
                    'Unit Display (Length)
                    '3: Inch
                    '0: mm
                    'FBase
                    '1 Decimal
                    '2 Fraction
                    'DRound
                    '0: Nearest Decimal
                    IntStatus = swDispDim.SetUnits2(False, 3, 2, 32, True, False)
                    'Error Check for set command
                    If IntStatus <> 0 Then
                        swApp.SendMsgToUser ("ERROR 1003: Failed to Set Units")
                        Exit Sub
                    End If
                End If
                'Get the next dimension
                Set swDispDim = swDispDim.GetNext5
            Wend

Next
'Rebuild document
swModel.ForceRebuild3 (True)
swModel.GraphicsRedraw2
End Sub

'Obsolete Comands From Other Rev

'Dim swDispDimA As Object
'Get Annotation from dimension
'Set swDispDimA = swDispDim.GetAnnotation()
'Set Style
'swDispDimA.SetStyleName "Fractional 32nd"
