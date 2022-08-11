Attribute VB_Name = "Samples"
Option Explicit

Public Sub YourProgramHere()
'This is a placeholder for your code
'You can uncomment any of the lines below and press F5
'to draw the corresponding shape or you can write your own program

butterfly
'Badge
'flower2
'draw_triangle
End Sub


Sub butterfly()
  Dim WingSize As Double, wingColors As Variant, size As Long, i As Long
  
  turtle.Reset
  turtle.DrawingMode = ttNoScreenRefresh
  turtle.TurnLeft 45
  
  wingColors = Array(ttred, ttblue, ttmagenta, ttyellow, ttgreen, ttgold)
  With turtle:
    .penColor = ttinvisible
    For size = 100 To 50 Step -10
      
      .fillColor = wingColors(size / 10 - 5)
      .PenDown
      For i = 1 To 4
        If i > 2 Then
          WingSize = size * 0.7
        Else
          WingSize = size
        End If
        .MoveCurved WingSize, WingSize / 4, ttPetalfd
        .MoveCurved -WingSize, WingSize / 4, ttPetalbk
        .TurnLeft 90
      Next i
      .PenUp
    Next size
    .PointInDirection 0
    .TurnRight 17
    .fillColor = ttinvisible
    .penColor = ttblack
    .PenDown
    .MoveCurved 100, 95, ttQuarterEllipse
    .PenUp
    .Move -100
    .TurnLeft 17 * 2
    .PenDown
    .MoveCurved 100, -95, ttQuarterEllipse
    .PenUp
    .Move -100
    
  End With
End Sub

Sub butterfly2()
  Dim WingSize As Double, wingColors As Variant, size As Long, i As Long
  
  turtle.Reset
  turtle.DrawingMode = ttNoScreenRefresh
  turtle.TurnLeft 45
  
  
  With turtle:
    .fillColor = ttcyan
    .penColor = ttinvisible
    For size = 100 To 5 Step -5
      
      .FillHueShift 17
      .PenDown
      For i = 1 To 4
        If i > 2 Then
          WingSize = size * 0.7
        Else
          WingSize = size
        End If
        .MoveCurved WingSize, WingSize / 4, ttPetalfd
        .MoveCurved -WingSize, WingSize / 4, ttPetalbk
        .TurnLeft 90
      Next i
      .PenUp
    Next size
    .PointInDirection 0
    .TurnRight 17
    .fillColor = ttinvisible
    .penColor = ttblack
    .PenDown
    .MoveCurved 100, 95, ttQuarterEllipse
    .PenUp
    .Move -100
    .TurnLeft 17 * 2
    .PenDown
    .MoveCurved 100, -95, ttQuarterEllipse
    .PenUp
    .Move -100
    .PointInDirection 90
  End With
End Sub

Sub Badge()
  Dim points As Long, i As Long, length As Long

  points = 6
  length = 100

  With turtle
    .Clear
    .PenUp
    .PointInDirection 0
    .FillType = ttSolid
    .fillColor = ttgold
    .penColor = ttinvisible
    .Center
    For i = 1 To points
      .Move length
      .Point
      .PenDown
      .Ellipse length / points
      .PenUp
      .Move -length
      .TurnRight 360 / points / 2
      .Move length * 0.6
      .Point
      .Move -length * 0.6
      .TurnRight 360 / points / 2
    Next i
    .ClosePoints
    .penColor = ttwhite
    .PenSize = length / points / 2
    .fillColor = ttinvisible
    .PointInDirection 90
    .PenDown
    .Ellipse 2 * length * 0.6 - length / points / 3
    .PenUp
    .FontSize = length / 3
    .FontColor = ttwhite
    .FontName = "Playbill"
    .WriteText "SHERIFF"
    .Group
    .fillColor = ttinvisible
    .penColor = ttblack
    .PenSize = 1
  End With

  turtle.GoToXY 20, 20
End Sub
  
  
  
Sub flower2()
  Dim j As Long, i As Long, sides As Long, petals As Long, color As ttColors

  sides = 6
  petals = 10
  color = ttorange
  
  With turtle
    .Center
    .Clear
    .fillColor = color
    .PenDown
    For j = 1 To petals
      .TurnRight 360 / petals
      For i = 1 To sides
        .MoveCurved 300 / sides, 110 / sides, ttHalfEllipse
        .TurnRight 360 / sides
      Next i
    Next j
    .PenUp
    .fillColor = ttinvisible
    .PenDown
  End With
End Sub

Sub flower1()
  Dim j As Long, i As Long, sides As Long, petals As Long, color As ttColors

  sides = 6
  petals = 10
  color = ttorange
  
  With turtle
    .Center
    .Clear
    .fillColor = color
    .PenDown
    For j = 1 To petals
      .TurnRight 360 / petals
      For i = 1 To sides
        .Move 300 / sides
        .TurnRight 360 / sides
      Next i
    Next j
    .PenUp
    .fillColor = ttinvisible
    .PenDown
  End With
End Sub



' from https://stackoverflow.com/questions/25772750/sierpinski-triangle-recursion-using-turtle-graphics
Sub sierpinski(length As Long, depth As Long)
  Dim i As Integer

  With turtle
    If depth = 0 Then
      For i = 0 To 2
        .Move length
        .TurnLeft 120
      Next i
      Exit Sub
    End If
    sierpinski length / 2, depth - 1
    .Move length / 2
    sierpinski length / 2, depth - 1
    .Move -length / 2
    .TurnLeft 60
    .Move length / 2
    .TurnRight 60
    sierpinski length / 2, depth - 1
    .TurnLeft 60
    .Move -length / 2
    .TurnRight 60
  End With
End Sub


Sub draw_triangle()
  Dim depth As Long
  
  depth = 3

  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    .FillType = ttSolid
    .fillColor = ttyellow
    .y = .y + 100
    .x = .x - 100
    sierpinski 200, depth
    .PenUp
    .fillColor = ttinvisible
  End With
End Sub

Sub star(ByVal points As Long, ByVal length As Double, ByVal interior_length)
  Dim i As Long
  With turtle
    For i = 1 To points
        .Move length
        .Point
        .Move -length
        .TurnRight 360 / points / 2
        .Move interior_length
        .Point
        .Move -interior_length
        .TurnRight 360 / points / 2
    Next i
    .ClosePoints
  End With
End Sub


Sub flower3()
  Dim sides As Long, color As ttColors, i As Long, layer As Long
  Dim length As Double, R1 As Double, R2 As Double
  
  sides = 20
  R2 = 200
  
  R1 = 20
  With turtle
    .Reset
    .PenUp
    .DrawingMode = ttNoScreenRefresh
    .fillColor = ttyellow
    .penColor = ttorange
    .PenSize = 3
    .FillTransparency = 0
    .PenTransparency = 0
    
    For layer = R2 + 150 To R2 Step -50
      length = .getSideLength(layer, sides)
      For i = 1 To sides
        .Move R1
        .PenDown
        .MoveCurved length, length ^ 2 / 500, ttPetalfd
        .MoveCurved -length, length ^ 2 / 500, ttPetalbk
        .PenUp
        .Move -R1
        .TurnRight 360 / sides
      Next i
      .TurnRight 360 / sides / 2
    Next layer
    .PenDown
    .PenSize = 6
    .Ellipse R1 * 2 + 3, R1 * 2 + 3
    .PenUp
  End With
End Sub


Sub batik_flower()
  Dim sides As Long, color As ttColors, i As Long
  Dim length As Double, R As Double, interior_sides As Long
  Dim interior_length As Double
  
  sides = 8
  R = 50
  interior_sides = 8
  
  
  interior_length = R / 1.5
  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    .PenUp
    .penColor = ttinvisible
    'Pistils
    .TurnLeft 90
    .Move R
    .TurnRight 45
    .fillColor = RGB(0, 128, 0)
    .PenDown
    length = .getSideLength(R, 4)
    For i = 1 To 4
      .MoveCurved R / 2, 0, ttLine
      .TurnRight 90
      .MoveCurved length, length, ttCusp
      .TurnRight 90
      .MoveCurved R / 2, 0, ttLine
      .TurnLeft 90
    Next i
    .PenUp
    
    'Petals
    length = .getSideLength(R, sides)
    .PointInDirection 90
    .Center
    .TurnLeft 360 / (2 * sides)
    .Move R
    .TurnRight 90 + 360 / (2 * sides)
    .fillColor = RGB(191, 191, 0)
    .PenDown
    For i = 1 To sides
     .MoveCurved length, length / 1.5, ttHalfEllipse
     .TurnRight 360 / sides
    Next i
    .PenUp
  
    
    .TurnLeft 90 + 360 / (2 * sides)
    .Move -R


    .fillColor = ttblack
    .TurnLeft 360 / 32
    star 16, R * 1.1, R * 1.1 * 0.6
    .TurnRight 360 / 32
    .fillColor = vbWhite
   
    For i = 1 To interior_sides
        .PenDown
        .MoveCurved interior_length, interior_length / 7, ttPetalfd
        .MoveCurved -interior_length, interior_length / 7, ttPetalbk
        .TurnLeft 360 / interior_sides
        .PenUp
    Next i
    .PenDown
    .fillColor = ttwhite
    .penColor = ttblack
    .PenSize = 3
    .Ellipse interior_length / 2.5
    .PenUp
    .Hide
  End With
  
  
End Sub

Sub Koch(depth As Long, length As Double)
  With turtle
    If depth = 1 Then
      .Move length
      Exit Sub
    End If
    Koch depth - 1, length / 3
    .TurnLeft 60
    Koch depth - 1, length / 3
    .TurnRight 120
    Koch depth - 1, length / 3
    .TurnLeft 60
    Koch depth - 1, length / 3
  End With
End Sub


Sub draw_snowflake()
  Dim i As Long, depth As Long, t As Single
  
  depth = 5
  t = Timer()
  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    .FillType = ttSolid
    .fillColor = ttyellow
    .PenUp
    .x = .x - 100
    .y = .y - 70
    .PenDown
    For i = 1 To 3
      Koch depth, 200
      .TurnRight 120
    Next i
    .PenUp
    .fillColor = ttinvisible
  End With
  Debug.Print Timer() - t
End Sub


Sub Spiral()
  ' Based on https://juliagraphics.github.io/Luxor.jl/v2.2/turtle/
Dim length, angle, d
Dim i

    d = 0.75
    
    length = 5
    angle = 89.5

    
    With turtle
      .Reset
      .Hide
      .CanvasColor = ttblack
      .DrawingMode = ttNoScreenRefresh
      .Hide
      .PenSize = 1
      .penColor = ttcyan
      For i = 1 To 400
        .PenDown
        .Move length
        .TurnRight angle
        length = length + d
        .PenHueShift 1
        .PenUp
      Next i
    End With
    
End Sub


Sub PolySpiral()

  Dim length, angle, d
  Dim c

    d = 1
    
    length = 300
    angle = 89

    
    With turtle
      .Reset
      .PenDown
      .PenSize = 0.5
      .fillColor = ttSkyBlue
      Do While length > d
        .Move length
        .TurnLeft angle
        length = length - d
      Loop
      .PenUp
      .Hide
    End With
    
End Sub

Sub Spiral2()

Dim length, angle, d
Dim i

    d = 0.75
    
    length = 5
    angle = 89.5

    
    With turtle
      .PenSize = 1
      .PenDown
      .Reset
      .DrawingMode = ttNoScreenRefresh
      .Hide
      For i = 1 To 400

        .Move length
        .TurnRight angle
        length = length + d
      Next i
      .PenUp
    End With
    
End Sub


Sub concentric()
  Dim circ As Long, segments As Long, diameter As Long, increment As Long, segment As Long
  Dim initialSegment As Double, angle As Double, newAngle As Double, levels As Long
  
  segments = 32
  diameter = 225
  increment = 20
  levels = 11
  
  With turtle
    .Reset
    For circ = 1 To levels
      
      initialSegment = Rnd() * 240 / segments - 120 / segments
      angle = initialSegment
      If circ <> levels Then
        .fillColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
      Else
        .fillColor = ttwhite
        segments = 1
      End If
      For segment = 1 To segments
        If segment <> segments Then
          newAngle = segment * 360 / segments + Rnd() * 240 / segments - 120 / segments
        Else
          newAngle = initialSegment
        End If
        If segments <> 1 Then
          .Arc diameter, diameter, angle, newAngle, ttsector
        Else
          .Arc diameter, diameter, angle, newAngle, TTARC
        End If
        angle = newAngle
        .FillHueShift 10
      Next segment
      diameter = diameter - increment
      segments = segments - 3
    Next circ
  
  End With
End Sub


'Inspired from Code a Pookkalam | Python programming https://www.youtube.com/watch?v=FYEuQUF37G0
Sub pookkalam1()
  Dim i As Long, j As Long, length As Double
  Dim white_circle_length As Double

  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    'external star
    .PenUp
    .fillColor = RGB(135, 16, 0)
    .penColor = ttinvisible
    star 24, 200, 190
    
    'orange circle filled with red
    .PenDown
    .penColor = RGB(250, 138, 49)
    .PenSize = 3
    .fillColor = RGB(236, 28, 0)
    .Ellipse 360, 360
    
    'dark orange star
    .fillColor = RGB(241, 93, 0)
    .penColor = ttinvisible
    .PenUp
    star 24, 160, 152
    
    'orange circle
    .fillColor = RGB(245, 169, 0)
    .PenDown
    .Ellipse 280, 280
    
    'yellow star
    .fillColor = RGB(248, 241, 0)
    .PenUp
    star 24, 120, 115
    
    ' 12 hexagons
    
    .penColor = RGB(136, 66, 0)
    .PointInDirection 90
    .PenDown
    length = .getSideLength(180, 6) / 2
    .fillColor = ttinvisible
    For j = 1 To 12
      For i = 1 To 6
        .Move length
        .TurnRight 360 / 6
      Next i
      .PenUp
      .TurnRight 360 / 12
      .PenDown
    Next j
  
    'white circle
    .Center
    .fillColor = ttwhite
    .penColor = ttinvisible
    .PenDown
    white_circle_length = (length - 15) * 2
    .Ellipse white_circle_length, white_circle_length
    
    'green star
    .PenUp
    length = (white_circle_length - 15) / 8
    .fillColor = RGB(8, 106, 0)
    star 24, length * 3, length * 3 - 3
    
    'orange internal circle
    .fillColor = RGB(245, 169, 0)
    .PenDown
    .Ellipse length * 4.5, length * 4.5
  
    'yellow internal circle
    .fillColor = RGB(248, 241, 0)
    .PenDown
    .Ellipse length * 3, length * 3
    
    'curved radials
    .PointInDirection 22.5
    
    Dim stepsForward As Double

    stepsForward = white_circle_length / 2 - 2.5
    .fillColor = ttinvisible
    .penColor = RGB(136, 66, 0)
    For i = 1 To 15
      .MoveBezier stepsForward, -45, 0.5 * stepsForward, 135, 0.7 * stepsForward
      .fillColor = .penColor
      .Ellipse 5
      .fillColor = ttinvisible
      .MoveBezier -stepsForward, -45, 0.5 * stepsForward, 135, 0.7 * stepsForward
      .TurnRight 360 / 15
    Next i
    .PenUp
  End With
End Sub


Sub mandala1()
  Dim i As Long, layer As Long, ShapePattern As clsShapeTransformer
  
  With turtle
    .Reset
    .fillColor = ttwhite
    .MoveCurved 200, 150, ttarccircle
    .MoveCurved -200, 150, ttarccircle
    'grab the shape to transform it
    Set ShapePattern = shapetransformer(.PenUp())
    For layer = 3 To 1 Step -1
      With ShapePattern
        If layer = 3 Or layer = 1 Then
          .Rotate 360 / 10 / 2
        End If
        If layer = 1 Then
          'clear central part
          turtle.PenDown
          turtle.penColor = ttinvisible
          turtle.Ellipse 100, 100
        End If
        For i = 1 To 5 - layer
          .Rotate 360 / 10, copyandrepeat:=10
          .Resize 0.85
        Next i
        .Resize 1.255
        If layer = 3 Then
          .Rotate -360 / 10 / 2
        End If
        .Translate -15
      End With
    Next layer
    With autoshapetransformer(msoShapeOval, 30, 20, ttblack)
      .Translate 145
      .Rotate -18
      .Rotate -36, copyandrepeat:=9
    End With
    
    .penColor = ttblack
    .Ellipse 32, 32
    .fillColor = ttblack
    .Ellipse 15, 15
    .Hide
  End With
End Sub
