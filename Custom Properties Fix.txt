' ******************************************************************************
' Created by Collin Brockman
' Gets Rid of Configuration Specific Custom Properties
' ******************************************************************************
'Preconditions:
'Open Part/Assembly
Dim swApp As Object
Dim PartDoc As SldWorks.ModelDoc2
Dim Part As SldWorks.ModelDoc2
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Public Sub main()
Set swApp = Application.SldWorks
Set PartDoc = swApp.ActiveDoc

'Check if the File type is a Assembly or Part
If PartDoc.GetType = swDocumentTypes_e.swDocDRAWING Then
    swApp.SendMsgToUser ("ERROR 1001: Wrong File Type")
    Exit Sub
End If

'Setup model as model extension object
Dim PartDocEx As SldWorks.ModelDocExtension
Set PartDocEx = PartDoc.Extension

'Setup list of needed properties for every part/assembly
Dim vValueSet(12) As Variant
vValueSet(0) = "PartNumber"
vValueSet(1) = "PartName"
vValueSet(2) = "PartName2"
vValueSet(3) = "Material"
vValueSet(4) = "DrawnBy"
vValueSet(5) = "Date"
vValueSet(6) = "HeatTreat"
vValueSet(7) = "Notes"
vValueSet(8) = "NextAssy"
vValueSet(9) = "Finish"
vValueSet(10) = "MachineOpt"
vValueSet(11) = "AngleOpt"
'vValueSet(10) = "Weight"

'Add Each Property in vValueSet to the custom properties
For j = 0 To ALen(vValueSet) - 1
    Dim CustPropMgr As SldWorks.CustomPropertyManager
    Set CustPropMgr = PartDocEx.CustomPropertyManager("")
    longstatus = CustPropMgr.Add3(vValueSet(j), 30, "", 0)
Next j
    

'Get Configuration Names
ReDim swConfigNames(PartDoc.GetConfigurationCount) As String
swConfigNames = PartDoc.GetConfigurationNames()

'For Each Config Do:
For i = 0 To ALen(swConfigNames) - 1

    'For Each Value in Custom Properties for Each Config Do:
    For j = 0 To ALen(vValueSet) - 1
        Dim CustPropMgr2 As SldWorks.CustomPropertyManager
        Set CustPropMgr2 = PartDocEx.CustomPropertyManager(swConfigNames(i))
        
        'Setup CustPropMgr2 As  Config Specific Properties
        Dim Val As String
        Dim EVal As String
        longstatus = CustPropMgr2.Get6(vValueSet(j), True, Val, EVal, boolstatus, False)
        
        'Setup CustPropMgr As Overall Properties
        Dim Val2 As String
        Dim EVal2 As String
        longstatus = CustPropMgr.Get6(vValueSet(j), True, Val2, EVal2, boolstatus, False)
        
        'Check there isnt a value in overall properties then set value from config specific to overall properties
        If Val2 = "" Then
            longstatus = CustPropMgr.Set2(vValueSet(j), Val)
        End If
    Next j

'Delete all Properties other than Weight in config specific
Dim vValueDel As Variant
vValueDel = CustPropMgr2.GetNames()
For k = 0 To ALen(vValueDel) - 1
    If vValueDel(k) <> "Weight" Then
        CustPropMgr2.Delete2 (vValueDel(k))
    End If
    
Next k

Next i

End Sub

'Function to Get Array Length
Public Function ALen(arr As Variant) As Integer
    ALen = UBound(arr) - LBound(arr) + 1
End Function
