' ******************************************************************************
' Created by Collin Brockman 7/16/2024
' Updated 10/28/2024
' Make Dimensions Toggle Parenthesis on Dimensions
' Gets any Selected Dimensions and Toggles the Selected Dimensions
' ******************************************************************************
'Start Function no Parameters in or out
Option Explicit
Dim boolstatus As Boolean
Sub main()
    'Initalize App and File to current instance
    Dim swApp As SldWorks.SldWorks
    Dim swModel As SldWorks.ModelDoc2
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    
    Dim swSelMgr As SldWorks.SelectionMgr
    Set swSelMgr = swModel.SelectionManager
    Dim i As Integer
    For i = 0 To swSelMgr.GetSelectedObjectCount2(-1)
        If swSelMgr.GetSelectedObjectType3(i, -1) = swSelectType_e.swSelDIMENSIONS Then
            'Set to Reference Dimension
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = swSelMgr.GetSelectedObject6(i, -1)
            If swDispDim.ShowParenthesis = False Then
                swDispDim.ShowParenthesis = True
            Else
                swDispDim.ShowParenthesis = False
        End If
    End If
    Next
    swModel.ClearSelection2 (True)
End Sub