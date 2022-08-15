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

Public Enum ttFlipH
  ttleft
  ttMidh
  ttRight
End Enum

Public Enum ttFlipV
  ttTop
  ttmidv
  ttBottom
End Enum

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
  x = myShape.left + myShape.width / 2
  y = myShape.top + myShape.height / 2
End Sub

Public Sub InitTransformer(shp As ShapeRange)
  Set myShape = shp
  myShapeName = shp.Name
  Set Canvas = ActiveSheet.ChartObjects("Canvas").Chart
  
End Sub

Public Function Rotate(angle As Double, Optional centerX, Optional centerY, Optional CopyAndRepeat As Long = 0)
  Dim R As Double, xc As Double, yc As Double, xo As Double, yo As Double
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
    R = Sqr((x - xc) ^ 2 + (y - yc) ^ 2)
    xo = R * Cos((angle * [pi()] / 180) + currentAngle)
    yo = R * Sin((angle * [pi()] / 180) + currentAngle)
    
    x = xc + xo
    y = yc + yo
    myShape.left = x - myShape.width / 2
    myShape.top = y - myShape.height / 2
    myShape.rotation = myShape.rotation + angle
    repetitions = repetitions + 1
  Loop Until repetitions >= CopyAndRepeat
  Set Rotate = Me
End Function

Public Function Translate(Optional x = 0, Optional y = 0, Optional steps = 0, Optional angle) As clsShapeTransformer
  Dim xo As Double, yo As Double

  myShape.left = myShape.left + x
  myShape.top = myShape.top + y
  If Not IsMissing(angle) Then
    xo = steps * Cos(angle * [pi()] / 180)
    yo = steps * Sin(angle * [pi()] / 180)
    myShape.left = myShape.left + xo
    myShape.top = myShape.top + yo

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
  myShape.ZOrder msoBringToFront
  Set Resize = Me
End Function

Public Function ZOrder(ByVal zordercmd As MsoZOrderCmd) As clsShapeTransformer
  myShape.ZOrder zordercmd
  Set ZOrder = Me
End Function

Public Function flipV(Optional axis As ttFlipV = ttmidv) As clsShapeTransformer
  myShape.Flip msoFlipVertical
  If axis = ttTop Then
    myShape.top = myShape.top - myShape.height
  ElseIf axis = ttBottom Then
    myShape.top = myShape.top + myShape.height
  End If
  Set flipV = Me
End Function

Public Function flipH(Optional axis As ttFlipH = ttMidh) As clsShapeTransformer
  myShape.Flip msoFlipHorizontal
  If axis = ttleft Then
    myShape.left = myShape.left - myShape.width
  ElseIf axis = ttRight Then
    myShape.left = myShape.left + myShape.width
  End If
  Set flipH = Me
End Function
Public Function Duplicate() As clsShapeTransformer
  Dim sr As ShapeRange

  Set sr = myShape.Duplicate
  sr.left = myShape.left
  sr.top = myShape.top
  Set Duplicate = Me
End Function


Public Function Center() As clsShapeTransformer
  myShape.left = Canvas.ChartArea.width / 2 - myShape.width / 2
  myShape.top = Canvas.ChartArea.height / 2 - myShape.height / 2
  Set Center = Me
End Function
