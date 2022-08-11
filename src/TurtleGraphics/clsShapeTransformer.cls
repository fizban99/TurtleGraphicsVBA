VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsShapeTransformer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private myShape As ShapeRange
Private myShapeName As String
Private Canvas As Chart

Private Sub getCanvasCenter(centerX, centerY, ByRef xc As Double, yc As Double)
  If Not IsMissing(centerX) Then
    xc = centerX
  Else
    xc = Canvas.ChartArea.width / 2
  End If
  If Not IsMissing(centerY) Then
    yc = centerY
  Else
    yc = Canvas.ChartArea.height / 2
  End If
End Sub

Private Sub getShapeCenter(ByRef x As Double, ByRef y As Double)
  x = myShape.Left + myShape.width / 2
  y = myShape.Top + myShape.height / 2
End Sub

Public Sub InitTransformer(shp As ShapeRange)
  Set myShape = shp
  myShapeName = shp.Name
  Set Canvas = ActiveSheet.ChartObjects("Canvas").Chart
End Sub

Public Function Rotate(angle As Double, Optional centerX, Optional centerY, Optional CopyAndRepeat As Long = 0)
  Dim r As Double, xc As Double, yc As Double, xo As Double, yo As Double
  Dim x As Double, y As Double, currentAngle As Double, sr As ShapeRange
  Dim repetitions As Long
  
  getCanvasCenter centerX, centerY, xc, yc
  getShapeCenter x, y
  Do
    If CopyAndRepeat > 0 Then
      Me.Duplicate
    End If
    If x - xc = 0 Then
      If y - yc < 0 Then
        currentAngle = -[pi()] / 2
      Else
        currentAngle = [pi()] / 2
      End If
    Else
      currentAngle = Atn((y - yc) / (x - xc))
      If x - xc < 0 Then
        currentAngle = currentAngle - [pi()]
      End If
    End If
    r = Sqr((x - xc) ^ 2 + (y - yc) ^ 2)
    xo = r * Cos((angle * [pi()] / 180) + currentAngle)
    yo = r * Sin((angle * [pi()] / 180) + currentAngle)
    
    x = xc + xo
    y = yc + yo
    myShape.Left = x - myShape.width / 2
    myShape.Top = y - myShape.height / 2
    myShape.Rotation = myShape.Rotation + angle
    repetitions = repetitions + 1
  Loop Until repetitions >= CopyAndRepeat
  Set Rotate = Me
End Function

Public Function Translate(Optional x = 0, Optional y = 0, Optional steps = 0, Optional angle) As clsShapeTransformer
  Dim xo As Double, yo As Double

  myShape.Left = myShape.Left + x
  myShape.Top = myShape.Top + y
  If Not IsMissing(angle) Then
    xo = steps * Cos(angle * [pi()] / 180)
    yo = steps * Sin(angle * [pi()] / 180)
    myShape.Left = myShape.Left + xo
    myShape.Top = myShape.Top + yo

  End If
  Set Translate = Me
End Function

Public Function Resize(ByVal factor As Single, Optional factorX, Optional factorY) As clsShapeTransformer
  If Not IsMissing(factorX) Then
    factor = factorX
  End If
  myShape.ScaleWidth factor, msoFalse, msoScaleFromMiddle
  If Not IsMissing(factorY) Then
    factor = factorY
  End If
  myShape.ScaleHeight factor, msoFalse, msoScaleFromMiddle
  Set Resize = Me
End Function

Public Function Duplicate() As clsShapeTransformer
  Dim sr As ShapeRange

  Set sr = myShape.Duplicate
  sr.Left = myShape.Left
  sr.Top = myShape.Top
  Set Duplicate = Me
End Function
