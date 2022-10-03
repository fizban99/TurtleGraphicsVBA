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


Public Function New_clsTurtle(Optional CanvasName = "Canvas", Optional shapeName = "Turtle") As clsTurtle
  Set New_clsTurtle = New clsTurtle
  New_clsTurtle.InitTurtle (CanvasName), (shapeName)
End Function


Public Function AutoShapeTransformer(ByVal autoShapeType As MsoAutoShapeType, _
  Optional ByVal width As Double = 100, Optional ByVal height As Double = 100, _
  Optional ByVal fillColor As ttColors = ttinvisible, Optional ByVal penColor As ttColors = ttblack, Optional ByVal penSize As Long = 1) _
   As clsShapeTransformer
  
  Static myShapeTransformer As New clsShapeTransformer
  Dim Canvas As Chart, ShapeToTransform As shape, X As Double, Y As Double
  Dim sr As ShapeRange
  
  Set Canvas = ActiveSheet.ChartObjects("Canvas").Chart
  X = Canvas.ChartArea.width / 2
  Y = Canvas.ChartArea.height / 2
  Set ShapeToTransform = Canvas.Shapes.AddShape(autoShapeType, X - width / 2, Y - height / 2, width, height)
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

Public Function ShapeTransformer(shape As Variant) As clsShapeTransformer
  Static myShapeTransformer As New clsShapeTransformer
  Dim Canvas As Chart, ShapeToTransform As ShapeRange, X As Double, Y As Double
  Dim sht As Worksheet, shp As shape
  
  Set Canvas = ActiveSheet.ChartObjects("Canvas").Chart
  If VarType(shape) = vbString Then
    Set sht = ActiveWorkbook.Worksheets("Shapes")
    Set shp = sht.Shapes(shape)
    shp.copy
    Canvas.Paste
    Set shp = Canvas.Shapes(Canvas.Shapes.Count)
    shp.left = Canvas.ChartArea.width / 2 - shp.width / 2
    shp.top = Canvas.ChartArea.height / 2 - shp.height / 2
    Set ShapeToTransform = Canvas.Shapes.Range(Array(Canvas.Shapes.Count))
  Else
    Set ShapeToTransform = shape
  End If
  X = Canvas.ChartArea.width / 2
  Y = Canvas.ChartArea.height / 2
  Set ShapeTransformer = myShapeTransformer
  myShapeTransformer.InitTransformer ShapeToTransform
End Function



