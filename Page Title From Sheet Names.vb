Dim swApp As Object
Public Sub main()
Set swApp = Application.SldWorks
'Opens Current Sheet
Dim swDrawing As DrawingDoc
Set swDrawing = swApp.ActiveDoc
Dim swSheet As Sheet
Set swSheet = swDrawing.GetCurrentSheet()
'Document Settings
Dim swSheetHeight As Double
swSheetHeight = 11 * 0.0254
Dim swSheetWidth As Double
swSheetWidth = 17 * 0.0254
Dim swSheetOff As Double
swSheetOff = 1 * 0.0254

Dim swSheetName As String
swSheetName = swSheet.GetName()

'Make Note Object
Dim myNote As SldWorks.Note
Set myNote = swDrawing.InsertNote("<FONT size=24PTS style=B effect=U><FONT style=I>" + swSheetName)
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
End Sub
