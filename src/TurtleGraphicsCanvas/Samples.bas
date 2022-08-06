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
  Dim WingSize As Single, wingColors As Variant, size As Long, i As Long
  
  turtle.Reset
  turtle.DrawingMode = ttNoScreenRefresh
  turtle.TurnLeft 45
  
  wingColors = Array(ttred, ttblue, ttmagenta, ttyellow, ttgreen, ttgold)
  With turtle:
    .PenColor = ttInvisible
    For size = 100 To 50 Step -10
      
      .FillColor = wingColors(size / 10 - 5)
      .Pendown
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
    .TurnRight 17
    .FillColor = ttInvisible
    .PenColor = ttBlack
    .Pendown
    .Movecurved 100, 95, ttQuarterEllipse
    .PenUp
    .Move -100
    .TurnLeft 17 * 2
    .Pendown
    .Movecurved 100, -95, ttQuarterEllipse
    .PenUp
    .Move -100
    
  End With
End Sub

Sub butterfly2()
  Dim WingSize As Single, wingColors As Variant, size As Long, i As Long
  
  turtle.Reset
  turtle.DrawingMode = ttNoScreenRefresh
  turtle.TurnLeft 45
  
  
  With turtle:
    .FillColor = ttcyan
    .PenColor = ttInvisible
    For size = 100 To 5 Step -5
      
      .FillHueShift 17
      .Pendown
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
    .TurnRight 17
    .FillColor = ttInvisible
    .PenColor = ttBlack
    .Pendown
    .Movecurved 100, 95, ttQuarterEllipse
    .PenUp
    .Move -100
    .TurnLeft 17 * 2
    .Pendown
    .Movecurved 100, -95, ttQuarterEllipse
    .PenUp
    .Move -100
    
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
    .FillColor = ttgold
    .PenColor = ttInvisible
    .Center
    For i = 1 To points
      .Move length
      .Point
      .Pendown
      .ellipse length / points
      .PenUp
      .Move -length
      .TurnRight 360 / points / 2
      .Move length * 0.6
      .Point
      .Move -length * 0.6
      .TurnRight 360 / points / 2
    Next i
    .ClosePoints
    .PenColor = ttwhite
    .PenSize = length / points / 2
    .FillColor = ttInvisible
    .PointInDirection 90
    .Pendown
    .ellipse 2 * length * 0.6 - length / points / 3
    .PenUp
    .FontSize = length / 3
    .FontColor = ttwhite
    .FontName = "Playbill"
    .WriteText "SHERIFF"
    .Group
    .FillColor = ttInvisible
    .PenColor = ttBlack
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
    .FillColor = color
    .Pendown
    For j = 1 To petals
      .TurnRight 360 / petals
      For i = 1 To sides
        .Movecurved 300 / sides, 110 / sides, ttHalfEllipse
        .TurnRight 360 / sides
      Next i
    Next j
    .PenUp
    .FillColor = ttInvisible
    .Pendown
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
    .FillColor = color
    .Pendown
    For j = 1 To petals
      .TurnRight 360 / petals
      For i = 1 To sides
        .Move 300 / sides
        .TurnRight 360 / sides
      Next i
    Next j
    .PenUp
    .FillColor = ttInvisible
    .Pendown
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
    .FillColor = ttyellow
    .y = .y + 100
    .x = .x - 100
    sierpinski 200, depth
    .PenUp
    .FillColor = ttInvisible
  End With
End Sub

Sub star(ByVal points As Long, ByVal length As Single)
  Dim i As Long
  With turtle
    For i = 1 To points
        .Move length
        .Point
        .Move -length
        .TurnRight 360 / points / 2
        .Move length * 0.6
        .Point
        .Move -length * 0.6
        .TurnRight 360 / points / 2
    Next i
    .ClosePoints
  End With
End Sub


Sub flower3()
  Dim sides As Long, color As ttColors, i As Long
  Dim length As Single, R As Single
  
  sides = 24
  length = 100
  
  R = 20
  With turtle
    .Reset
    .PenUp
    .FillColor = ttyellow
    For i = 1 To sides
      .Move R
      .Pendown
      .Movecurved length, length / 5, ttPetalbk
      .Movecurved -length, length / 5, ttPetalfd
      .PenUp
      .Move -R
      .TurnRight 360 / sides
    Next i
  End With
End Sub


Sub batik_flower()
  Dim sides As Long, color As ttColors, i As Long
  Dim length As Single, R As Single, interior_sides As Long
  Dim interior_length As Single
  
  sides = 8
  R = 50
  interior_sides = 8
  
  
  interior_length = R / 1.5
  With turtle
    .Reset
    .PenUp
    .PenColor = ttInvisible
    'Pistils
    .TurnLeft 90
    .Move R
    .TurnRight 45
    .FillColor = RGB(0, 128, 0)
    .Pendown
    length = R * (2 * Sin([Pi()] / 4))
    For i = 1 To 4
      .Movecurved R / 2, 0, ttLine
      .TurnRight 90
      .Movecurved length, length, ttCusp
      .TurnRight 90
      .Movecurved R / 2, 0, ttLine
      .TurnLeft 90
    Next i
    .PenUp
    
    'Petals
    length = R * (2 * Sin([Pi()] / sides))
    .PointInDirection 90
    .Center
    .TurnLeft 360 / (2 * sides)
    .Move R
    .TurnRight 90 + 360 / (2 * sides)
    .FillColor = RGB(191, 191, 0)
    .Pendown
    For i = 1 To sides
     .Movecurved length, length / 1.5, ttHalfEllipse
     .TurnRight 360 / sides
    Next i
    .PenUp
  
    
    .TurnLeft 90 + 360 / (2 * sides)
    .Move -R


    .FillColor = ttBlack
    .TurnLeft 360 / 32
    star 16, R * 1.1
    .TurnRight 360 / 32
    .FillColor = vbWhite
   
    For i = 1 To interior_sides
        .Pendown
        .Movecurved interior_length, interior_length / 7, ttPetalfd
        .Movecurved -interior_length, interior_length / 7, ttPetalbk
        .TurnLeft 360 / interior_sides
        .PenUp
    Next i
    .Pendown
    .FillColor = ttwhite
    .PenColor = ttBlack
    .PenSize = 3
    .ellipse interior_length / 2.5
    .PenUp
    .Hide
  End With
  
  
End Sub

Sub Koch(depth As Long, length As Single)
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
  Dim i As Long, depth As Long
  
  depth = 5
  
  With turtle
    .Reset
    .FillType = ttSolid
    .FillColor = ttyellow
    .PenUp
    .x = .x - 100
    .y = .y - 70
    .Pendown
    For i = 1 To 3
      Koch depth, 200
      .TurnRight 120
    Next i
    .PenUp
    .FillColor = ttInvisible
  End With
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
      .CanvasColor = ttBlack
      .DrawingMode = ttNoScreenRefresh
      .Hide
      .PenSize = 1
      .PenColor = ttcyan
      For i = 1 To 400
        .Pendown
        .Move length
        .TurnRight angle
        length = length + d
        .PenHueShift 1
        .PenUp
      Next i
      .Hide
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
      .Pendown
      .PenSize = 0.5
      .FillColor = ttSkyBlue
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
      .Pendown
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
  Dim initialSegment As Single, angle As Single, newAngle As Single, levels As Long
  
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
        .FillColor = RGB(Rnd() * 255, Rnd() * 255, Rnd() * 255)
      Else
        .FillColor = ttwhite
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



