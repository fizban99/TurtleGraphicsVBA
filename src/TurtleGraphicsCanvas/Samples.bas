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
      
      .fillcolor = wingColors(size / 10 - 5)
      .PenDown
      For i = 1 To 4
        If i > 2 Then
          WingSize = size * 0.7
        Else
          WingSize = size
        End If
        .Movecurved WingSize, WingSize / 4, ttPetalfd
        .Movecurved -WingSize, WingSize / 4, ttPetalbk
        .TurnLeft 90
      Next i
      .PenUp
    Next size
    .PointInDirection 0
    .turnright 17
    .fillcolor = ttinvisible
    .penColor = ttblack
    .PenDown
    .Movecurved 100, 95, ttQuarterEllipse
    .PenUp
    .Move -100
    .TurnLeft 17 * 2
    .PenDown
    .Movecurved 100, -95, ttQuarterEllipse
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
    .fillcolor = ttcyan
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
        .Movecurved WingSize, WingSize / 4, ttPetalfd
        .Movecurved -WingSize, WingSize / 4, ttPetalbk
        .TurnLeft 90
      Next i
      .PenUp
    Next size
    .PointInDirection 0
    .turnright 17
    .fillcolor = ttinvisible
    .penColor = ttblack
    .PenDown
    .Movecurved 100, 95, ttQuarterEllipse
    .PenUp
    .Move -100
    .TurnLeft 17 * 2
    .PenDown
    .Movecurved 100, -95, ttQuarterEllipse
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
    .fillcolor = ttgold
    .penColor = ttinvisible
    .Center
    For i = 1 To points
      .Move length
      .Point
      .PenDown
      .Ellipse length / points
      .PenUp
      .Move -length
      .turnright 360 / points / 2
      .Move length * 0.6
      .Point
      .Move -length * 0.6
      .turnright 360 / points / 2
    Next i
    .ClosePoints
    .penColor = ttwhite
    .penSize = length / points / 2
    .fillcolor = ttinvisible
    .PointInDirection 90
    .PenDown
    .Ellipse 2 * length * 0.6 - length / points / 3
    .PenUp
    .FontSize = length / 3
    .FontColor = ttwhite
    .FontName = "Playbill"
    .WriteText "SHERIFF"
    .Group
    .fillcolor = ttinvisible
    .penColor = ttblack
    .penSize = 1
  End With

  turtle.gotoxy 20, 20
End Sub
  
  
  
Sub flower_arc()
  Dim j As Long, i As Long, sides As Long, petals As Long, color As ttColors

  sides = 16
  petals = 16
  color = ttblack
  
  With turtle
    .Reset
    .Clear
    .fillcolor = color
    
    .PenDown
    For j = 1 To petals
      .turnright 360 / petals
      For i = 1 To sides
        .Movecurved 500 / sides, 250 / sides, ttarccircle
        .turnright 360 / sides
      Next i
    Next j
    .PenUp
    .fillcolor = ttinvisible
    .PenDown
  End With
End Sub

Sub flower_ellipse()
  Dim j As Long, i As Long, sides As Long, petals As Long, color As ttColors

  sides = 16
  petals = 16
  color = ttblack
  
  With turtle
    .Reset
    .Clear
    .fillcolor = color
    .PenDown
    For j = 1 To petals
      .turnright 360 / petals
      For i = 1 To sides
        .Movecurved 500 / sides, 200 / sides, ttHalfEllipse
        .turnright 360 / sides
      Next i
    Next j
    .PenUp
    .fillcolor = ttinvisible
    .PenDown
  End With
End Sub


Sub flower1()
  Dim j As Long, i As Long, sides As Long, petals As Long, color As ttColors

  sides = 16
  petals = 16
  color = ttblack
  
  With turtle
    .Reset
    .Clear
    .fillcolor = color
    .PenDown
    For j = 1 To petals
      .turnright 360 / petals
      For i = 1 To sides
        .Move 500 / sides
        .turnright 360 / sides
      Next i
    Next j
    .PenUp
    .fillcolor = ttinvisible
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
    .turnright 60
    sierpinski length / 2, depth - 1
    .TurnLeft 60
    .Move -length / 2
    .turnright 60
  End With
End Sub


Sub sierpinski_triangle()
  Dim depth As Long
  
  depth = 4

  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    .FillType = ttSolid
    .fillcolor = ttblack
    .y = .y + 100
    .x = .x - 100
    sierpinski 200, depth
    .PenUp
    .fillcolor = ttinvisible
  End With
End Sub

Sub star(ByVal points As Long, ByVal length As Double, ByVal interior_length)
  Dim i As Long
  With turtle
    For i = 1 To points
        .Move length
        .Point
        .Move -length
        .turnright 360 / points / 2
        .Move interior_length
        .Point
        .Move -interior_length
        .turnright 360 / points / 2
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
    .fillcolor = ttyellow
    .penColor = ttorange
    .penSize = 3
    .FillTransparency = 0
    .PenTransparency = 0
    
    For layer = R2 + 150 To R2 Step -50
      length = .getSideLength(layer, sides)
      For i = 1 To sides
        .Move R1
        .PenDown
        .Movecurved length, length ^ 2 / 500, ttPetalfd
        .Movecurved -length, length ^ 2 / 500, ttPetalbk
        .PenUp
        .Move -R1
        .turnright 360 / sides
      Next i
      .turnright 360 / sides / 2
    Next layer
    .PenDown
    .penSize = 6
    .Ellipse R1 * 2 + 3, R1 * 2 + 3
    .PenUp
  End With
End Sub


Sub circle_checkered()
  ' inspired by https://sp.depositphotos.com/99992684/stock-illustration-monochrome-elegant-pattern-black-and.html
  Dim sides As Long, color As ttColors, i As Long, length As Double
  Dim radius As Double
  
  sides = 24
  radius = 250
  
  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    .penSize = 0.5
    .penColor = ttinvisible
    .fillcolor = ttblack
    'place a black circle as background
    'to invert background and foreground
    .Ellipse radius * 2 + 1, radius * 2 + 1
    .fillcolor = ttwhite
    For i = 1 To sides
      .Movecurved radius, radius / 2, ttarccircle
      .Movecurved -radius, radius / 2, ttarccircle
      .turnright 360 / sides
    Next i
    .PenUp
  End With
End Sub

Sub batik_flower()
  Dim sides As Long, color As ttColors, i As Long
  Dim length As Double, r As Double, interior_sides As Long
  Dim interior_length As Double
  
  sides = 8
  r = 50
  interior_sides = 8
  
  
  interior_length = r / 1.5
  With turtle
    .Reset
    .DrawingMode = ttNoScreenRefresh
    .PenUp
    .penColor = ttinvisible
    'Pistils
    .TurnLeft 90
    .Move r
    .turnright 45
    .fillcolor = RGB(0, 128, 0)
    .PenDown
    length = .getSideLength(r, 4)
    For i = 1 To 4
      .Movecurved r / 2, 0, ttLine
      .turnright 90
      .Movecurved length, length, ttCusp
      .turnright 90
      .Movecurved r / 2, 0, ttLine
      .TurnLeft 90
    Next i
    .PenUp
    
    'Petals
    length = .getSideLength(r, sides)
    .PointInDirection 90
    .Center
    .TurnLeft 360 / (2 * sides)
    .Move r
    .turnright 90 + 360 / (2 * sides)
    .fillcolor = RGB(191, 191, 0)
    .PenDown
    For i = 1 To sides
     .Movecurved length, length / 1.5, ttHalfEllipse
     .turnright 360 / sides
    Next i
    .PenUp
  
    
    .TurnLeft 90 + 360 / (2 * sides)
    .Move -r


    .fillcolor = ttblack
    .TurnLeft 360 / 32
    star 16, r * 1.1, r * 1.1 * 0.6
    .turnright 360 / 32
    .fillcolor = vbWhite
   
    For i = 1 To interior_sides
        .PenDown
        .Movecurved interior_length, interior_length / 7, ttPetalfd
        .Movecurved -interior_length, interior_length / 7, ttPetalbk
        .TurnLeft 360 / interior_sides
        .PenUp
    Next i
    .PenDown
    .fillcolor = ttwhite
    .penColor = ttblack
    .penSize = 3
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
    .turnright 120
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
    .fillcolor = ttblack
    .Ellipse 250, 250
    .fillcolor = ttwhite
    .PenUp
    .x = .x - 100
    .y = .y - 70
    .PenDown
    For i = 1 To 3
      Koch depth, 200
      .turnright 120
    Next i
    shapetransformer(.PenUp()).Center
    shapetransformer(.Group()).MergeShapes msoMergeSubtract
    
    .fillcolor = ttinvisible
  End With
  Debug.Print Timer() - t
End Sub


Sub colored_polyspiral()
  ' Based on https://juliagraphics.github.io/Luxor.jl/v2.2/turtle/
Dim length, angle, d
Dim i

    d = 0.75
    
    length = 5
    angle = 89.5

    
    With turtle
      .Reset
      autoshapetransformer msoShapeRectangle, 420, 420, ttblack, ttblack
      .Hide
      .DrawingMode = ttNoScreenRefresh
      .Hide
      .penSize = 1
      .penColor = ttcyan
      For i = 1 To 400
        .PenDown
        .Move length
        .turnright angle
        length = length + d
        .penhueshift 1
        .PenUp
      Next i
      
    End With
    
End Sub


Sub RainbowSpiral()
  ' Based on https://docs.racket-lang.org/racket_turtle/racket_turtle_examples_with_recursion.html
Dim length As Long, angle As Double, d As Double
Dim i As Long, colors As Variant

    d = 0.75
    
    length = 5
    angle = 59.5
    colors = Array(ttred, ttgreen, ttyellow, ttplum, ttblue, ttorange)
    
    With turtle
      .Reset
      .Hide
      autoshapetransformer(msoShapeRectangle, 630, 630, ttblack).Translate 0, 10
      .DrawingMode = ttNoScreenRefresh
      .Hide
      .penSize = 2
      .penColor = ttcyan
      For i = 1 To 300
        .PenDown
        .Move length
        .turnright angle
        length = length + d
        .penColor = colors(i Mod 6)
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
      .DrawingMode = ttNoScreenRefresh
      .PenDown
      .penSize = 0.5
      .fillcolor = ttblack
      Do While length > d
        .Move length
        .TurnLeft angle
        length = length - d
      Loop
      shapetransformer(.PenUp()).Center
      .Hide
      .PenDown
      .Center
      .Ellipse 300, 300
      shapetransformer(.Group()).MergeShapes msoMergeIntersect
    End With
    
End Sub

Sub Spiral2()

Dim length, angle, d
Dim i

    d = 0.75
    
    length = 5
    angle = 89.5

    
    With turtle
      .penSize = 1
      .PenDown
      .Reset
      .DrawingMode = ttNoScreenRefresh
      .Hide
      For i = 1 To 400

        .Move length
        .turnright angle
        length = length + d
      Next i
      .PenUp
      .Center
      .PenDown
      .fillcolor = ttwhite
      .penColor = ttinvisible
      .Ellipse 500, 500
      .Ellipse 300, 300
      
      shapetransformer(.Group(2)).MergeShapes msoMergeSubtract
      .fillcolor = ttinvisible
      .penSize = 4
      .penColor = ttblack
      .Ellipse 300, 300
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
        .fillcolor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
      Else
        .fillcolor = ttwhite
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
    .fillcolor = RGB(135, 16, 0)
    .penColor = ttinvisible
    star 24, 200, 190
    
    'orange circle filled with red
    .PenDown
    .penColor = RGB(250, 138, 49)
    .penSize = 3
    .fillcolor = RGB(236, 28, 0)
    .Ellipse 360, 360
    
    'dark orange star
    .fillcolor = RGB(241, 93, 0)
    .penColor = ttinvisible
    .PenUp
    star 24, 160, 152
    
    'orange circle
    .fillcolor = RGB(245, 169, 0)
    .PenDown
    .Ellipse 280, 280
    
    'yellow star
    .fillcolor = RGB(248, 241, 0)
    .PenUp
    star 24, 120, 115
    
    ' 12 hexagons
    
    .penColor = RGB(136, 66, 0)
    .PointInDirection 90
    .PenDown
    length = .getSideLength(180, 6) / 2
    .fillcolor = ttinvisible
    For j = 1 To 12
      For i = 1 To 6
        .Move length
        .turnright 360 / 6
      Next i
      .PenUp
      .turnright 360 / 12
      .PenDown
    Next j
  
    'white circle
    .Center
    .fillcolor = ttwhite
    .penColor = ttinvisible
    .PenDown
    white_circle_length = (length - 15) * 2
    .Ellipse white_circle_length, white_circle_length
    
    'green star
    .PenUp
    length = (white_circle_length - 15) / 8
    .fillcolor = RGB(8, 106, 0)
    star 24, length * 3, length * 3 - 3
    
    'orange internal circle
    .fillcolor = RGB(245, 169, 0)
    .PenDown
    .Ellipse length * 4.5, length * 4.5
  
    'yellow internal circle
    .fillcolor = RGB(248, 241, 0)
    .PenDown
    .Ellipse length * 3, length * 3
    
    'curved radials
    .PointInDirection 22.5
    
    Dim stepsForward As Double

    stepsForward = white_circle_length / 2 - 2.5
    .fillcolor = ttinvisible
    .penColor = RGB(136, 66, 0)
    For i = 1 To 15
      .MoveBezier stepsForward, -45, 0.5 * stepsForward, 135, 0.7 * stepsForward
      .fillcolor = .penColor
      .Ellipse 5
      .fillcolor = ttinvisible
      .MoveBezier -stepsForward, -45, 0.5 * stepsForward, 135, 0.7 * stepsForward
      .turnright 360 / 15
    Next i
    .PenUp
  End With
End Sub


Sub mandala1()
  Dim i As Long, layer As Long, ShapePattern As clsShapeTransformer
  
  With turtle
    .Reset
    .fillcolor = ttwhite
    .Movecurved 200, 150, ttarccircle
    .Movecurved -200, 150, ttarccircle
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
    .fillcolor = ttblack
    .Ellipse 15, 15
    .Hide
  End With
End Sub


Sub overlapped_flower()
  Dim i As Long

  With turtle
    .Reset
    .PenUp
    .fillcolor = ttblack
    .penColor = ttwhite
    .penSize = 2
    For i = 1 To 8
    .PenDown
    .Movecurved 100, 70, ttarccircle
    .Movecurved -100, 70, ttarccircle
    .turnright 360 / 8
    .PenUp
    Next i
    .PenDown
    .Center
    .PointInDirection 90
    .Movecurved 100, 70, ttarccircle
    .PenUp
    .Hide
  End With
End Sub


Sub framed_hexagon()
    Dim t As Double, i As Long, sr As Shape, vert As Variant
  
    With turtle
      .Reset
      .fillcolor = 12632256
      .TurnLeft 90
      .Move 100
      .TurnLeft 120
      .Move 50
      .TurnLeft 60
      .Move 100
      .TurnLeft 60
      .Move 100
      .TurnLeft 60
      .Move 50
      .TurnLeft 120
      .Move 100
      t = 2 * Sqr(50 ^ 2 - 25 ^ 2)
      shapetransformer(.PenUp()).Translate(-t).Rotate 360 / 6, copyandrepeat:=5
      
    End With
End Sub



Sub intertwined_rect_substract()
  Dim r As Double, d As Double, sides As Long
  Dim side1 As Double, side2 As Double
  
  r = 150
  d = 15
  sides = 11
  With turtle
    side1 = .getSideLength(r, sides / 2)
    side2 = .getSideLength(r + d, sides / 2)
    .Reset
    .JoinType = ttRoundJoin
    .PenUp
    'side to cut
    .TurnLeft 90
    .Move r + d
    .turnright 90 + 360 / sides
    .PenDown
    .penColor = ttblack
    .fillcolor = ttgold
    .Move side2
    .turnright 360 / sides * 2
    .Move side2
    .turnright 90 + 360 / sides
    .Move d
    .turnright 90 - 360 / sides
    .Move side1
    .TurnLeft 360 / sides * 2
    .Move side1
    .turnright 90 - 360 / sides
    .PenUp
    'first cutter
    .Center
    .PointInDirection 360 / sides
    .Move r + d
    .turnright 90 + 360 / sides
    .PenDown
    .Move side2 / 2
    .turnright 90 + 360 / side2 * 2
    .Move side1
    .turnright 90
    .Move side2
    .turnright 90
    .Move side1
    .PenUp
    shapetransformer(.Group()).MergeShapes (msoMergeSubtract)
    'second cutter
    .Center
    .PointInDirection 360 / sides * 3
    .Move r
    .turnright 90 + 360 / sides
    .PenDown
    .Move side1
    .TurnLeft 90
    .Move side1
    .TurnLeft 90
    .Move side1
    .PenUp
    shapetransformer(.Group()).MergeShapes (msoMergeSubtract)
    shapetransformer(.Group()).Rotate 360 / sides * 2, copyandrepeat:=10
  End With
  

End Sub


Sub intertwinded_rect()
  Dim side1 As Double, side2 As Double, i As Long
  Dim r As Double, j As Long, d As Double, start_x As Double, start_y As Double
  Dim firstX As Double, firstY As Double, sides As Long
  Dim shift As Double, sideb As Double, length As Double
  
  r = 150
  d = 15
  sides = 11
  With turtle
    .Reset

    .PenUp
    side1 = .getSideLength(r, sides / 2)
    side2 = .getSideLength(r + d, sides / 2)
    
    .JoinType = ttRoundJoin
    .TurnLeft 90
    .Move r
    .turnright 90 + 360 / sides
    ' draw orange background
    .PenDown
    .penColor = ttgold
    .fillcolor = ttgold
    length = side1
    For j = 1 To 2
      For i = 1 To sides
          .Move length
          .turnright 360 / 11 * 2
      Next i
      If j = 1 Then
        .TurnLeft 90 + 360 / sides
        .Move d
        .turnright 90 + 360 / sides
        length = side2
      End If

    Next j
    .PenUp
    .PointInDirection 90
    .Center
    .TurnLeft 90 + 360 / sides
    .Move r
    .penSize = 1
    .PointInDirection 90
    .penColor = ttblack

    .fillcolor = ttinvisible
    .PenDown

    length = side1
    ' draw black background
    For j = 1 To 2
      For i = 1 To sides
          .Move length
          .turnright 360 / 11 * 2
      Next i
      .PenUp
      .TurnLeft 90 + 360 / sides
      .Move d
      .PointInDirection 90
      length = side2
      .PenDown
      .penColor = ttblack
    Next j
    .PenUp

    .PointInDirection 90
    .Center
    .TurnLeft 90 + 360 / sides
    .Move r
    firstX = .x
    firstY = .y
    .PointInDirection 90
    .fillcolor = ttwhite
    ' draw twice
    ' the first one the lines up
    ' the second one the lines down
    ' since they overlap the lines up
    ' a shift at the beginning is required for the background
    ' (points)
    ' the same shift overlaps the lines
    .JoinType = ttRoundJoin
    .CapType = ttflat
    shift = 5
    For j = 1 To 2
      For i = 1 To sides
        start_x = .x
        start_y = .y
        If j = 2 Then
          .PenDown
          .Move shift
          .Point
        Else
          .Move shift
        End If
        .Move side1 / 2 - shift * 2

        If j = 2 Then
          .Point
          .Move shift
          .PenUp
        Else
          .Move shift
          .Point
          .PenDown
        End If
        .Move side1 / 2 - shift
        If j = 1 Then
          .Point
          .Move shift

          .PenUp
        Else
          .Move shift
        End If
        .TurnLeft 90 - 360 / sides
        .Move d
        .TurnLeft 90 + 360 / sides
        If j = 1 Then
          .PenDown
          .Move shift
          .Point
        Else
          .Move shift
        End If
        .Move side2 / 2 - shift
        
        If j = 1 Then
          .Point
          .PenUp
        Else
          .PenDown
          .Move shift
          .Point
          .Move side2 / 2 - shift * 2
          .Point
          .Move shift
          .PenUp
        End If
        .penColor = ttgold
        .fillcolor = ttgold
        .ClosePoints SendToBack:=2
        .Group 3, ungroupfirst:=False
        .penColor = ttblack
        .gotoxy start_x, start_y
        .turnright 180
        .Move side1
        .turnright 360 / sides * 2
        .PenUp
      Next i
      .gotoxy firstX, firstY
      .PointInDirection 90
    Next j
    
  End With
End Sub


Sub interwinded_rect_single()
  Dim side1 As Double, side2 As Double, i As Long, hypotenuse1 As Double, hypotenuse2 As Double
  Dim r As Double, j As Long, d As Double, angle As Double, start_x As Double, start_y As Double
  Dim factor As Long
  
  r = 100
  d = 10
  angle = 22.5
  With turtle
    .Reset

    .PenUp
    side1 = .getSideLength(r, 11 / 2)
    side2 = .getSideLength(r + d, 11 / 2)
    .TurnLeft 90 + 360 / 11
    .Move r
    .PointInDirection 90
    .fillcolor = ttwhite
    factor = 1

      For i = 1 To 11
        start_x = .x
        start_y = .y
        .PenDown
        .Point
        .Move side1
        .Point
        .PenUp
        .TurnLeft factor * (90 - 360 / 11)
        .Move d
        .TurnLeft factor * (90 + 360 / 11)
        .PenDown
        .Point
        '.MoveCurved hypotenuse2, -hypotenuse2, ttarccircle
        .Move side2
        .Point
        .PenUp
        .penColor = ttinvisible
        .ClosePoints SendToBack:=2
        .penColor = ttblack
        .gotoxy start_x, start_y
        .turnright 180
        .Move side1
        .turnright factor * (360 / 11 * 2)
      Next i

    
  End With
End Sub




Sub intertwinded_star()
  Dim side1 As Double, side2 As Double, i As Long, angle As Double
  Dim r As Double, j As Long, d As Double, start_x As Double, start_y As Double
  Dim firstX As Double, firstY As Double, sides As Long
  Dim hypotenuse1 As Double, hypotenuse2 As Double
  
  r = 150
  d = 50
  sides = 11
  
  angle = 22.5

  With turtle
    .Reset
    .penColor = ttblack
    .PenUp
    side1 = .getSideLength(r, sides / 2)
    side2 = .getSideLength(r + d, sides / 2)
    hypotenuse1 = side1 / Cos(angle / 180 * [pi()]) / 2
    hypotenuse2 = side2 / Cos(angle / 180 * [pi()]) / 2
    .TurnLeft 90 + 360 / sides
    .Move r
    .PointInDirection 90
    .fillcolor = ttgold
    firstX = .x
    firstY = .y
    'Repeat twice
    'Each time draws half the sides
    For j = 1 To 2
      For i = 1 To sides
        start_x = .x
        start_y = .y
        If j = 2 Then
          .PenDown
          .Point
        End If
        .turnright angle
        .Move hypotenuse1
        .TurnLeft angle
        .Point
        If j = 2 Then
          .PenUp
        Else
          .PenDown
        End If
        .TurnLeft angle
        .Move hypotenuse1
        .turnright angle
        If j = 1 Then
          .PenUp
          .Point
        End If
        .TurnLeft 90 - 360 / sides
        .Move d
        .TurnLeft 90 + 360 / sides
        If j = 1 Then
          .PenDown
          .Point
        End If
        .TurnLeft angle
        .Move hypotenuse2
        .turnright angle
        .Point
        If j = 1 Then
          .PenUp
        Else
          .PenDown
          .turnright angle
          .Move hypotenuse2
          .TurnLeft angle
          .PenUp
          .Point
        End If

        .penColor = ttgold
        .penSize = 0
        .fillcolor = ttgold
        .ClosePoints SendToBack:=2

        .Group 3, ungroupfirst:=False
        .penColor = ttblack
        .fillcolor = ttinvisible
        .penSize = 1
        .gotoxy start_x, start_y
        .turnright 180
        .Move side1
        .turnright 360 / sides * 2
        .PenUp
      Next i
      .gotoxy firstX, firstY
      .PointInDirection 90
    Next j
    
  End With
End Sub



Sub intertwinded_star_single()
  Dim side1 As Double, side2 As Double, i As Long, angle As Double
  Dim r As Double, j As Long, d As Double, start_x As Double, start_y As Double
  Dim firstX As Double, firstY As Double, sides As Long
  Dim hypotenuse1 As Double, hypotenuse2 As Double
  
  r = 147
  d = 56
  sides = 11
  
  angle = 22.5

  With turtle
    .Reset
    .penColor = ttgold
    side1 = .getSideLength(r + d / 2, sides / 2)
    'side2 = .getSideLength(r + d, sides / 2)
    hypotenuse1 = side1 / Cos(angle / 180 * [pi()]) / 2
    .TurnLeft 90 + 360 / sides
    .PenUp
    .Move r + d / 2
    .penSize = d / 2
    .PointInDirection 90
    .fillcolor = ttinvisible
    .PenDown
      For i = 1 To sides
        .turnright angle
        .Move hypotenuse1
        .TurnLeft angle
        .TurnLeft angle
        .Move hypotenuse1
        .turnright angle
        .turnright 360 / sides * 2
      Next i
      .PenUp
      .gotoxy firstX, firstY
      .PointInDirection 90

    
  End With
End Sub


Sub intertwinded_curved_transform()
  Dim side1 As Double, side2 As Double, i As Long, angle As Double
  Dim r As Double, j As Long, d As Double
  Dim firstX As Double, firstY As Double, sides As Long
  Dim hypotenuse1 As Double, hypotenuse2 As Double
  Dim radius1 As Double, radius2 As Double, disp As Double
  Dim arcx_st As Double, arcy_st As Double, arcx_end As Double, arcy_end As Double
  
  r = 150
  d = 25
  
  sides = 11
  
  angle = 22.5

  With turtle
    .Reset
    .penColor = ttblack
    .fillcolor = ttinvisible
    .PenUp
    side1 = .getSideLength(r, sides / 2)
    radius1 = .getSideLength(r, 7 / 2) / 2
    disp = d * Cos(180 / 11 / [pi()])
    radius2 = radius1 + d
    hypotenuse1 = side1 / Cos(angle / 180 * [pi()]) / 2
    .TurnLeft 90
    .Move r
    firstX = .x
    firstY = .y
    .Move disp
     arcx_end = .x
     arcy_end = .y
    .Move -disp
    .PointInDirection 90 + 360 / sides
    For i = 1 To 2
      .PenDown
      .turnright angle
      .Movecurved hypotenuse1, radius1, ttarccircle
      .TurnLeft angle
      If i = 1 Then
        .PenUp
      End If
      .TurnLeft 90
      .Movecurved disp, 0, ttLine
      If i = 1 Then
        .PenDown
      End If
      .Movexycurved arcx_end, arcy_end, -radius2, ttarccircle
      .turnright angle
      shapetransformer(.PenUp()).ZOrder msoSendToBack
      .gotoxy firstX, firstY
      .PointInDirection 90 + 360 / sides
      .penColor = ttgold
      .penSize = 0
      .fillcolor = ttgold
    Next i
    With shapetransformer(.Group())
      .Rotate 360 / 11, copyandrepeat:=11
      .flipH ttleft
      .Rotate 360 / 11, copyandrepeat:=11
    End With
    .Hide
  End With
End Sub

Sub squared_knot()
  Dim i As Long

  With turtle
    .Reset
    .fillcolor = ttgold
    .Move 100
    .TurnLeft 90
    .Move 25
    .TurnLeft 90
    .Move 75
    For i = 1 To 3
      .turnright 90
      .Move 25
    Next i
    .TurnLeft 90
    .Move 25
    .TurnLeft 90
    .Move 50
    .TurnLeft 90
    .Move 75
    .TurnLeft 90
    .Move 75
    

    With shapetransformer(.PenUp())
        .Duplicate.spin 90
        .Translate 87.5, 12.5

        .Duplicate.spin 90
        .Translate 37.5, 37.5
    End With
    With shapetransformer(.Group())

        .Duplicate.flipV().flipH
        .Translate y:=100
    End With
    .Center
    .PointInDirection 180
    .PenDown
    For i = 1 To 2
      .Move 75
      .TurnLeft 90
      .Move 25
      .TurnLeft 90
    Next i
    With shapetransformer(.PenUp())
      .Translate x:=50
      .Duplicate.Translate x:=100
      .Duplicate.spin 90
      .Translate x:=-50
    End With
    With shapetransformer(.Group())
      .spin 45
    End With
  End With
    
    
    
End Sub


Sub fat_star()
  Dim side1 As Double, side2 As Double, i As Long, angle As Double
  Dim r As Double, j As Long, d As Double, start_x As Double, start_y As Double
  Dim firstX As Double, firstY As Double, sides As Long
  Dim hypotenuse1 As Double, hypotenuse2 As Double
  
  r = 200
  d = 20
  sides = 11
  
  angle = 45

  With turtle
    .Reset
    .JoinType = ttRoundJoin
    .PenUp
    side1 = .getSideLength(r, sides)
    hypotenuse1 = side1 / Cos(angle / 180 * [pi()]) / 2
    .TurnLeft 90 + 360 / sides / 2
    .Move r
    .PointInDirection 90
    .fillcolor = ttblack
    firstX = .x
    firstY = .y
    'Repeat twice
    .penSize = d
    For j = 1 To 2
      For i = 1 To sides
        .PenDown
        start_x = .x
        start_y = .y
        .turnright angle
        .Movecurved hypotenuse1, hypotenuse1 * 1.5, ttarccircle
        .TurnLeft angle
        .TurnLeft angle
        .Movecurved hypotenuse1, hypotenuse1 * 1.5, ttarccircle
        .turnright angle
        .turnright 360 / sides
      Next i
      .PenUp
      .gotoxy firstX, firstY
      .PointInDirection 90
      .penColor = ttwhite
      .penSize = d - 2
    Next j
    With shapetransformer(.Group()).Duplicate.Resize(0.5).Duplicate.Resize(0.25)
    autoshapetransformer msoShapeRectangle, 450, 450, ttinvisible, ttinvisible
    End With
  End With
End Sub

Sub fat_star2()
  Dim side1 As Double, side2 As Double, i As Long, angle As Double
  Dim r As Double, j As Long, d As Double, start_x As Double, start_y As Double
  Dim firstX As Double, firstY As Double, sides As Long
  Dim hypotenuse1 As Double, hypotenuse2 As Double
  
  r = 200
  d = 10
  sides = 11
  
  angle = 45

  With turtle
    .Reset

    .PenUp
    side1 = .getSideLength(r, sides / 2)
    hypotenuse1 = side1 / Cos(angle / 180 * [pi()]) / 2
    .TurnLeft 90 + 360 / sides / 2
    .Move r
    .PointInDirection 90
    .fillcolor = ttblack
    .penColor = ttblack
    firstX = .x
    firstY = .y
    'Repeat twice
    .penSize = d
    For j = 1 To 2
      For i = 1 To sides
        .PenDown
        start_x = .x
        start_y = .y
        .turnright angle
        .Movecurved hypotenuse1, hypotenuse1 * 1.5, ttarccircle
        .TurnLeft angle
        .TurnLeft angle
        .Movecurved hypotenuse1, hypotenuse1 * 1.5, ttarccircle
        .turnright angle
        .turnright 360 / sides * 2
      Next i
      .PenUp
      .gotoxy firstX, firstY
      .PointInDirection 90
      .penColor = ttwhite
      .penSize = d - 2
    Next j
    With shapetransformer(.Group()).Center().Duplicate.Resize(0.5).Duplicate.Resize(0.25)
    autoshapetransformer msoShapeRectangle, 450, 450, ttinvisible, ttblack
    End With
  End With
End Sub



Sub triquetra_arc(side_ext As Double, side_int As Double, r As Double)
  Dim width As Single
  
  width = (side_ext - side_int) / 2
  With turtle
    .Center
    .PointInDirection 0
    .PenUp
    .Move r
    .PenDown
    .turnright 90 + 60
    .Movecurved side_ext, -side_ext / 2, ttHalfEllipse
    .TurnLeft 180
    .PenUp
    .Move width
    .PenDown
    .Movecurved side_int, side_int / 2, ttHalfEllipse
    'add a margin to avoid a hanging line
    .turnright 360 / 3
    .Move width
    .PenUp
    With shapetransformer(.Group(2))
      .MergeShapes msoMergeSubtract
    End With
  End With


End Sub

Sub triquetra()
  Dim side_ext As Double, width As Single, side_int As Double, r As Double
  
  r = 100
  
  width = 20
  With turtle
    .Reset
    .fillcolor = ttgold
    side_ext = .getSideLength(r, 3)
    side_int = side_ext - width * 2
    triquetra_arc side_ext, side_int, r
    .Center
    .PointInDirection 180
    .Move side_ext * Sqr(3) / 2 - r
    .PointInDirection 0
    .PenDown
    .Ellipse side_ext, side_ext
    .PenUp
    shapetransformer(.Group()).MergeShapes msoMergeSubtract
    .PenDown
    'draw an arc for the second time
    triquetra_arc side_ext, side_int, r
    
    .Center
    .PointInDirection 360 / 3 * 2
    .Move r
    .PointInDirection 30
    .Move width
    .PenDown
    .turnright 45
    .Movecurved side_int / Sqr(2), -side_int / 2, ttarccircle
    .PointInDirection 90
    .Move side_ext / 2
    .PointInDirection 180
    .Move side_ext / 2
    .PenUp
    shapetransformer(.Group(2)).MergeShapes msoMergeSubtract
    shapetransformer(.Group(1)).flipH(ttRight).Translate x:=-width
    With shapetransformer(.Group())
      .MergeShapes (msoMergeUnion)

      .Rotate 120, copyandrepeat:=2
    End With

    
  End With
  
End Sub

Sub pookkalam2()
  Dim i As Integer, fillcolor As Variant, side As Variant, r As Double
  With turtle
    .Reset
    .penColor = ttblack
    .fillcolor = ttblack
    r = .getRadiusLength(200, 4)
    .Ellipse r * 2, r * 2
    .fillcolor = 246
    fillcolor = .fillcolor
    side = Array(200, 185, 173, 162, 150)

    For i = 0 To 4
      With autoshapetransformer(msoShapeRectangle, width:=side(i), Height:=side(i), fillcolor:=fillcolor, penColor:=ttinvisible)
        .Rotate 360 / 80 * i
        .spin 360 / 40, 40
        turtle.FillHueShift 20
        If i < 3 Then
          fillcolor = turtle.fillcolor
        Else
          fillcolor = ttwhite
        End If
      End With
    Next i
    r = .getRadiusLength(130, 4)
    .penSize = 4
    .penColor = ttbrown
    .fillcolor = ttplum
    .Ellipse r * 2, r * 2
    .PenUp
    .fillcolor = ttgold
    For i = 0 To 3
      star 24, r - 4 - i * 10, r - 14 - i * 10
      .FillHueShift 20
    Next i
    .PenDown
    .penhueshift 30
    .Ellipse (r - 14 - 3.5 * 10) * 2, (r - 14 - 3.5 * 10) * 2
    .PenUp
    .penSize = 2

    For i = 5 To 7
      star 12, r - 4 - i * 10, r - 14 - i * 10
      .FillHueShift 30
    Next i
    .Hide
  End With
  
End Sub

Sub phyllotaxis()
  Dim angle As Double, r As Double, c As Double, i As Long
  
  c = 12
  
  
  With turtle
    .Reset
    .Hide
    autoshapetransformer msoShapeRectangle, 550, 550, ttblack
    .fillcolor = ttyellow
    .penColor = ttinvisible
    .DrawingMode = ttScreenRefresh
    For i = 0 To 500
      angle = i * 137.508
      r = c * Sqr(i)
      .PointInDirection angle
      .FillHueShift 1
      .PenUp
      .Move r
      .PenDown
      .Ellipse 15, 15
      .PenUp
      .Move -r
    Next i
  
  End With

End Sub


Sub chussy_mandala()
  Dim g As ShapeRange, side As Variant, s As Double, fillcolor As Long, i As Long
  
  
  With turtle
  .Reset
  .penColor = ttbrown
  .fillcolor = ttgold
  .penSize = 5
  .Ellipse 500, 500

  fillcolor = 246
  s = .getSideLength(250, 4)
  side = Array(s - 4, s - 30, s - 60, s - 90, s - 120)
  

  End With
  
    For i = 0 To 4
      With autoshapetransformer(msoShapeRectangle, width:=side(i), Height:=side(i), fillcolor:=fillcolor, penColor:=ttinvisible)
        .Rotate 360 / 80 * i
        .spin 360 / 40, 40
        turtle.FillHueShift 20
        If i < 3 Then
          fillcolor = turtle.fillcolor
        Else
          fillcolor = ttgold
        End If
      End With
    Next i

  With shapetransformer("chussy")
    .Translate y:=-170
    .Rotate 36, copyandrepeat:=10
    .Translate steps:=60, angle:=90
    .Resize 0.8
    .Rotate 18
    .Rotate 36, copyandrepeat:=10
    .Translate steps:=100, angle:=90
    .Resize 0.8
    .Rotate 18
    .Rotate 36, copyandrepeat:=10
    .Center
    .Resize 0.8
    .Rotate -18
  End With
  With turtle
    .fillcolor = ttinvisible
    .Ellipse 120, 120
    .Hide
  End With
End Sub
