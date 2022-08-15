Attribute VB_Name = "ClassProvider"
Option Explicit

Public Function Turtle() As clsTurtle
  Static init As Boolean
  Static myTurtle As New clsTurtle
  
  Set Turtle = myTurtle
  If init = False Or myTurtle.ImageLost Then
    myTurtle.InitTurtle "Canvas", "Turtle"
    init = True
  End If
End Function


Public Function New_clsTurtle(Optional CanvasName = "Canvas", Optional ShapeName = "Turtle") As clsTurtle
  Set New_clsTurtle = New clsTurtle
  New_clsTurtle.InitTurtle (CanvasName), (ShapeName)
End Function


Public Function AutoShapeTransformer(autoShapeType As MsoAutoShapeType, _
  Optional width As Double = 100, Optional height As Double = 100, _
  Optional fillColor As ttColors = ttinvisible, Optional penColor As ttColors = ttblack, Optional penSize As Long = 1) _
   As clsShapeTransformer
  
  Static myShapeTransformer As New clsShapeTransformer
  Dim Canvas As Chart, ShapeToTransform As Shape, x As Double, y As Double
  Dim sr As ShapeRange
  
  Set Canvas = ActiveSheet.ChartObjects("Canvas").Chart
  x = Canvas.ChartArea.width / 2
  y = Canvas.ChartArea.height / 2
  Set ShapeToTransform = Canvas.Shapes.AddShape(autoShapeType, x - width / 2, y - height / 2, width, height)
  With ShapeToTransform
    .line.Weight = penSize
    If fillColor <> ttinvisible Then
      .Fill.ForeColor.RGB = fillColor
    Else
      .Fill.Visible = msoFalse
    End If
    If penColor <> ttinvisible Then
       .line.ForeColor.RGB = penColor
    Else
       .line.Visible = msoFalse
    End If
    
  End With
  Set sr = Canvas.Shapes.Range(Array(ShapeToTransform.Name))
  Set AutoShapeTransformer = myShapeTransformer
  myShapeTransformer.InitTransformer sr
End Function

Public Function ShapeTransformer(sr As ShapeRange) As clsShapeTransformer
  Static myShapeTransformer As New clsShapeTransformer
  Dim Canvas As Chart, ShapeToTransform As ShapeRange, x As Double, y As Double
  
  Set Canvas = ActiveSheet.ChartObjects("Canvas").Chart
  Set ShapeToTransform = sr
  x = Canvas.ChartArea.width / 2
  y = Canvas.ChartArea.height / 2
  Set ShapeTransformer = myShapeTransformer
  myShapeTransformer.InitTransformer ShapeToTransform
End Function



