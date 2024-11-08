' ******************************************************************************
' Created by Collin Brockman
' ******************************************************************************
'Start Function no Parameters in or out
Option Explicit
Dim boolstatus As Boolean
Dim longstatus As Long
Sub main()
    'Initalize App and File to current instance
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    'Setup model as model extension object
    Dim PartDocEx As SldWorks.ModelDocExtension
    Set PartDocEx = swModel.Extension
    
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

    Dim CustPropMgr2 As SldWorks.CustomPropertyManager
    Set CustPropMgr2 = PartDocEx.CustomPropertyManager("")
        
    'Setup CustPropMgr2 As  Config Specific Properties
    Dim Val As String
    Dim EVal As String
    longstatus = CustPropMgr2.Get6("ERNumber", True, Val, EVal, boolstatus, False)
    
    'Create REV Symbol as a Note uses settings from definition
    Dim REVSym As SldWorks.Note
    Set REVSym = swModel.InsertNote("<TL-" + Val + ">")
    

End Sub
