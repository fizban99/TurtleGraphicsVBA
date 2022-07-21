# TurtleGraphicsVBA
Turtle Graphics using VBA in Excel. An add-in created for the purpose of teaching basic programming to young students. The advantage of using Excel is that if Excel is already on the computer, no additional installation is required besides these two files. The graphics generated are in vector format. They can be exported to emf and converted to svg using inkscape, for example or online through https://convertio.co/emf-svg/.


https://user-images.githubusercontent.com/13463439/179361846-f0ffedd4-76ff-4b72-a0cb-e866c45c2b1f.mp4

Note that the TurtleGraphics.xlam file is locked for visualization so that VBA errors made by students do not jump to lines in that module. To actually edit or view the module, the password is "turtle".

## Commands

### `Arc(DiameterAcross As Double, DiameterFrontBack As Variant, StartAngle As Double, EndAngle As Double, ArcType As ttArcType) `

### `Center()`

### `Clear()`

### `ClosePoints()`

### `Ellipse(DiameterAcross As Double, Optional DiameterFrontBack)`

### `GoToXY(ByVal X As Long, ByVal Y As Long)`

### `Group()`

### `Hide()`

### `Move(ByVal steps As Double)`

### `PenDown()`

### `PenUp()`

### `Point()`

### `PointInDirection(ByVal angle As Integer)`

### `SaveCanvas(fileName As String, ImageFormat As ttImageFormat)`

### `Show()`

### `TurnLeft(ByVal angle As Double)`

### `TurnRight(ByVal angle As Double)`

### `Wait(milliseconds As Long)`

### `WriteText(txt As String)`


## Properties

### `CanvasColor`

### `CanvasHeight` (Read Only)

### `CanvasWidth` (Read Only)

### `DrawingMode`
  Whether the shape is drawn upon pen up (to speed up the drawing) or while it is being drawn.

### `FillColor`

### `FillType`

### `FontColor`

### `FontName`

### `FontSize`

### `FontStyle`

### `FontWeight`

### `Height` (Read Only)

### `LineStyle`

### `PenColor`

### `PenSize`

### `Width` (Read Only)

### `X`

### `Y`
