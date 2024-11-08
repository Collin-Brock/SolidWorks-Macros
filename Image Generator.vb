' ******************************************************************************
' Created by Collin Brockman
' Generate an Image from selected view
' ******************************************************************************
'Preconditions:
'Open Drawing
Option Explicit
Dim swApp As SldWorks.SldWorks
Dim SelMgr As Object
Dim swFeatMgr As SldWorks.FeatureManager

' Define two types
Type DoubleRec
    dValue As Double
End Type

Type Int2Rec
    iLower As Long ' Assuming that a C int has 4 bytes
    iUpper As Long
End Type

Sub main()
     'Setup current drawing and sheet
    Set swApp = Application.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Dim swDraw As SldWorks.DrawingDoc
    Dim boolstatus As Boolean
    Dim swModelEx As ModelDocExtension
    Set swModel = swApp.ActiveDoc
    'Set swDraw = swModel
    Set swModelEx = swModel.Extension
    
    'Check if the instance is running
    If swModel Is Nothing Then
        swApp.SendMsgToUser ("ERROR 1002: No Open File")
        Exit Sub
    End If
    
    swModel.ViewZoomtofit2
    boolstatus = swModelEx.SetUserPreferenceToggle(swUserPreferenceToggle_e.swImageQualityApplyToAllReferencedPartDoc, swUserPreferenceOption_e.swDetailingNoOptionSpecified, True)
    swModel.SetTessellationQuality (106)
    swModel.GraphicsRedraw2
    Set swFeatMgr = swModel.FeatureManager
    Dim vLightProp As Variant
    
    Dim i2r As Int2Rec
    Dim dr As DoubleRec
    
    ' Forcibly turn lights off
    i2r.iLower = 1
    i2r.iUpper = 1
    LSet dr = i2r
    
    Dim i As Integer
    For i = 0 To swModel.GetLightSourceCount - 1

        If InStr(swModel.GetLightSourceName(i), "Ambient") > 0 Then
    
        vLightProp = swModel.LightSourcePropertyValues(i)
        vLightProp(15) = 0.75
        swModel.LightSourcePropertyValues(i) = vLightProp
        Else
        vLightProp = swModel.LightSourcePropertyValues(i)
        vLightProp(17) = dr.dValue
        swModel.LightSourcePropertyValues(i) = vLightProp
        End If
    Next i
    swFeatMgr.UpdateFeatureTree
    swModel.GraphicsRedraw2
    
     'Create Path Name
    Dim Path As String
    Path = swModel.GetPathName
    Dim PathArray() As String
    PathArray = Split(Path, ".")
    Path = PathArray(0)
    PathArray = Split(Path, "\")
    Path = PathArray(UBound(PathArray))
    Dim Path2 As String
    Path2 = "P:\Engineering Images (Macro)\" + Path + "-png" + ".png"
    Dim Path3 As String
    Path3 = "P:\Engineering Images (Macro)\" + Path + "-jpg" + ".jpg"
    'Save As
    '2 save as a copy
    Dim LongError As Long
    Dim LongWarning As Long
    boolstatus = swModelEx.SaveAs3(Path2, 0, 2, Nothing, Nothing, LongError, LongWarning)
    boolstatus = swModelEx.SaveAs3(Path3, 0, 2, Nothing, Nothing, LongError, LongWarning)
    
    
    boolstatus = swModelEx.SetUserPreferenceInteger(swUserPreferenceIntegerValue_e.swImageQualityShaded, swUserPreferenceOption_e.swDetailingNoOptionSpecified, swImageQualityShaded_e.swShadedImageQualityFine)
    
End Sub


