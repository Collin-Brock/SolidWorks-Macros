' ******************************************************************************
' Created By Collin Brockman
' ******************************************************************************
Dim swApp As Object

Dim PartDoc As SldWorks.ModelDoc2
Dim Part As SldWorks.ModelDoc2
Dim boolstatus As Boolean
Dim longstatus As Long, longwarnings As Long
Public Grind As Boolean
Public ViewFront As Boolean
Public Sub main()
Set swApp = Application.SldWorks
Set PartDoc = swApp.ActiveDoc
Dim Path As String
Path = PartDoc.GetPathName

'Get Configuration Names
ReDim swConfigNames(PartDoc.GetConfigurationCount) As String
swConfigNames = PartDoc.GetConfigurationNames()

'Get Part Name from Path
Dim PathArray() As String
PathArray = Split(Path, "\")
PathArray = Split(PathArray(ALen(PathArray) - 1), ".")
Dim PartName As String
PartName = PathArray(0)

'Document Settings
Dim swSheetHeight As Double
swSheetHeight = 11 * 0.0254
Dim swSheetWidth As Double
swSheetWidth = 17 * 0.0254
Dim swSheetOff As Double
swSheetOff = 1 * 0.0254
Dim ScaleNum As Integer
ScaleNum = 1
Dim ScaleDen As Integer
ScaleDen = 4

'Create New Document
Set Part = swApp.NewDocument(Path, 12, swSheetWidth, swSheetHeight)
'Opens Current Sheet
Dim swDrawing As DrawingDoc
Set swDrawing = Part
Dim swSheet As sheet
Set swSheet = swDrawing.GetCurrentSheet()
Dim vPropSheet As Variant
vPropSheet = swSheet.GetProperties2
swSheet.SetProperties2 vPropSheet(0), vPropSheet(1), ScaleNum, ScaleDen, False, swSheetWidth, swSheetHeigth, vPropSheet(7)
For i = 0 To ALen(swConfigNames) - 1
    If Not i = 0 Then
        boolstatus = Part.NewSheet3(swConfigNames(i), 12, 12, ScaleNum, ScaleDen, False, Path, swSheetWidth, swSheetHeight, "Default")
    End If
    swSheet.SetName (swConfigNames(i))
    Dim myView As SldWorks.View
    PartDoc.ShowConfiguration2 (swConfigNames(i))
    swSheet.ReloadTemplate (False)
    
    'Make Note Object
    Dim myNote As SldWorks.Note
    Set myNote = Part.InsertNote("<FONT size=24PTS style=B effect=U><FONT style=I>" + swConfigNames(i))
    'Sets Note to Origin (Must be done to ensure reliability)
    myNote.SetTextPoint 0, 0, 0
    'Sets Note Position to upper right hand corner based on the upper right corner of the note
    Dim TextPos() As Double
    TextPos = myNote.GetUpperRight
    'Set X Position
    TextPos(0) = swSheetWidth - 0.5 * swSheetOff - TextPos(0)
    'Set Y Position
    TextPos(1) = swSheetHeight - swSheetOff - TextPos(1)
    myNote.SetTextPoint TextPos(0), TextPos(1), 0
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'swDrawing.Create3rdAngleViews2 (Path)
    '
    'boolstatus = swDrawing.GenerateViewPaletteViews(Path)
    'Set myView = swDrawing.DropDrawingViewFromPalette2("*Current", swSheetWidth * 0.25, swSheetHeight / 2, 0)
    '
    Dim ViewDir As String
    ViewDir = "*Front"
    Set myView = Part.CreateDrawViewFromModelView3(Path, ViewDir, swSheetWidth * 0.35, swSheetHeight / 2, 0)
    ViewDir = "*Right"
    Set myView = swDrawing.CreateUnfoldedViewAt3(swSheetWidth * 0.85, swSheetHeight / 2, 0, False)
    'Set myView = Part.CreateDrawViewFromModelView3(Path, ViewDir, swSheetWidth * 0.85, swSheetHeight / 2, 0)
Next i

Set PartDoc = swApp.ActiveDoc
Dim vPropSet(6) As Variant
vPropSet(0) = "ERNumber"
vPropSet(1) = "ERNotes"
vPropSet(2) = "ERDate"
vPropSet(3) = "CheckedBy"
vPropSet(4) = "CheckedByDate"
vPropSet(5) = "ERBy"

Dim MyDate As Date
MyDate = Date
Dim ShortDate As String
ShortDate = Format(MyDate, "m/d/yy")
Dim vValueSet(6) As Variant
vValueSet(0) = "-"
vValueSet(1) = "INITIAL RELEASE"
vValueSet(2) = MyDate
vValueSet(3) = " "
vValueSet(4) = ShortDate
vValueSet(5) = "CDB"

'Setup model as model extension object
Dim PartDocEx As SldWorks.ModelDocExtension
Set PartDocEx = PartDoc.Extension

For j = 0 To ALen(vPropSet) - 1
    Dim CustPropMgr As SldWorks.CustomPropertyManager
    Set CustPropMgr = PartDocEx.CustomPropertyManager("")
    longstatus = CustPropMgr.Set2(vPropSet(j), vValueSet(j))
Next j
PartDoc.ForceRebuild3 (True)
End Sub
Public Function ALen(arr As Variant) As Integer
    ALen = UBound(arr) - LBound(arr) + 1
End Function
    'swDrawing.Create3rdAngleViews2 (Path)
    'boolstatus = swDrawing.GenerateViewPaletteViews(Path)
    'Set myView = swDrawing.DropDrawingViewFromPalette2("*Current", swSheetWidth * 0.25, swSheetHeight / 2, 0)