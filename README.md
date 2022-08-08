# TurtleGraphicsVBA
Turtle Graphics using VBA in Excel. An add-in created for the purpose of teaching basic programming to young students. The advantage of using Excel is that if Excel is already on the computer, no additional installation is required besides these two files. The graphics generated are in vector format. They can easily copied and resized in any Office application and they can be exported to emf and converted to svg using Inkscape, for example.

Besides the normal straight line movement, this library has a moveCurved function and moveBezier function that allow moving through a curved path, producing with low effort much visually pleasant shapes than traditional turtle-graphics shapes, allowing the creation of mandalas, batik patterns and pookkalams.

Between a penDown and a penUp instruction, the move instruction will create a polyline (all segments will belong to the same shape). Similarly, between a penDown and a penUp instruction, the moveCurved will produce a Bezier curve, but without having to worry about the control points. Both methods cannot be mixed, though. If you need a straight line within a moveCurved path, use ttLine as CurveType for moveCurved third parameter. 

Note that the TurtleGraphics.xlam file is locked for visualization so that VBA errors made by students do not jump to lines in that module. To actually edit or view the module, the password is "turtle".

You can see some sample drawings in the [samples](https://github.com/fizban99/TurtleGraphicsVBA/tree/main/samples) folder.

![Main screen](./img/main-screen.png?raw=true)


## Commands

### `Arc(ByVal DiameterAcross As Double, ByVal DiameterFrontBack As Variant, ByVal StartAngle As Double, ByVal EndAngle As Double, ByVal ArcType As ttArcType) `

### `Center()`

### `Clear()`

### `ClosePoints()`

### `Ellipse(ByVal DiameterAcross As Double, Optional DiameterFrontBack)`

### `GoToXY(ByVal X As Long, ByVal Y As Long)`

### `Group()`

### `Hide()`

### `Move(ByVal steps As Double)`

### `MoveBezier(ByVal Steps As Double, angle1 As Double, ByVal distance1 As Double, ByVal angle2 As Double, ByVal distance2 As Double)`

### `Sub MoveCurved(ByVal Steps As Double, ByVal radius As Double, Optional curveType As ttCurveType = ttHalfEllipse)`

### `PenDown()`

### `PenUp()`

### `Point()`

### `PointInDirection(ByVal angle As Integer)`

### `SaveCanvas(ByVal fileName As String, ByVal ImageFormat As ttImageFormat)`

### `Show()`

### `TurnLeft(ByVal angle As Double)`

### `TurnRight(ByVal angle As Double)`

### `Wait(ByVal milliseconds As Long)`

### `WriteText(ByVal txt As String)`


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

### `LineStyle`

### `PenColor`

### `PenSize`

### `X`

### `Y`
