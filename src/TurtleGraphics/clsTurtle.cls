VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTurtle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'**************************
'Class clsTurtle
'Based on an idea of Robin M Snyder (http://ccscjournal.willmitchell.info/Vol13-97/east97/Robin%20M%20Snyder-2.pdf)
'**************************
Option Explicit

'To save as emf (see http://stackoverflow.com/questions/1791369/excel-export-chart-to-wmf-or-emf)

Private Declare Function OpenClipboard _
    Lib "user32" ( _
        ByVal hWnd As Long) _
As Long

Private Declare Function CloseClipboard Lib "user32" () As Long

Private Declare Function GetClipboardData _
    Lib "user32" ( _
        ByVal wFormat As Long) _
As Long

Private Declare Function EmptyClipboard Lib "user32" () As Long

'// CreateMetaFileA DeleteEnhMetaFile
Private Declare Function CopyEnhMetaFileA _
    Lib "gdi32" ( _
        ByVal hENHSrc As Long, _
        ByVal lpszFile As String) _
As Long

Private Declare Function DeleteEnhMetaFile _
    Lib "gdi32" ( _
        ByVal hemf As Long) _
As Long


Private Declare PtrSafe Sub ColorRGBToHLS Lib "shlwapi.dll" (ByVal clrRGB As Long, _
    wHue As Integer, wLuminance As Integer, wSaturation As Integer)

Private Declare Function ColorHLSToRGB Lib "shlwapi.dll" _
                                      (ByVal wHue As Long, _
                                       ByVal wLuminance As Long, _
                                       ByVal wSaturation As Long) As Long
Const PI = 3.14159265358979
Const INITIAL_ANGLE = 90
Public Enum ttColors
  ttwhite = vbWhite
  ttblack = vbBlack
  ttyellow = vbYellow
  ttblue = vbBlue
  ttmagenta = vbMagenta
  ttcyan = vbCyan
  ttred = vbRed
  ttgreen = vbGreen
  ttgold = 55295
  ttorange = 52479
  ttbrown = 13209
  ttAqua = 13421619
  ttSkyBlue = 16763904
  ttViolet = 8388736
  ttRose = 13408767
  ttplum = 6697881
  ttinvisible = -1
End Enum

Public Enum ttImageFormat
  ttGIF
  ttJPG
  ttEMF
  ttPNG
  ttPDF
End Enum

Public Enum ttCurveType
  ttarccircle
  ttHalfEllipse
  ttQuarterEllipse
  ttPetalfd
  ttPetalbk
  ttCusp
  ttLine
End Enum

Public Enum ttCurveEnding
  ttNormal
  ttFlatstart
  ttFlatend
  ttFlatBoth
End Enum

Public Enum ttCapType
  ttroundcap
  ttflat
  ttsquare
End Enum

Public Enum ttJoinType
  ttRoundJoin
  ttBevel
  ttMiter
End Enum

Public Enum ttArcType
  TTARC = 0
  ttSegment = 1
  ttsector = 2
End Enum

Public Enum ttLineStyle
  ttSolidLine = 0
End Enum

Public Enum ttDrawingMode
  ttScreenRefresh = 1
  ttNoScreenRefresh = 2
End Enum
 
Public Enum ttFontStyle
  ttFSNormal = 0
  ttfsItalic = 1
  
End Enum
Public Enum ttFontWeight
  ttFWNormal = 0
  ttFWBold = 1
End Enum
 
 
Public Enum ttFillType
  ttSolid = 1
End Enum

Dim PenX As Double
Dim PenY As Double
Dim InitialPenX As Double
Dim InitialPenY As Double
Dim TurtleImage As shape
Dim Canvas As Chart
Dim CanvasShape As shape
Dim IsPenDown As Boolean
Dim nAngle As Double

Dim PolyPoints As FreeformBuilder
Dim tempPoly As shape
Dim RefreshLines As Long
Dim SAoP() As Variant
Dim PolyNodes As Long
Dim IsCurvedPoly As Boolean
Dim BufLimit As Long
Dim NumLines As Long
Dim penHue As Integer, penLuminosity As Integer, penSaturation As Integer
Dim fillHue As Integer, fillLuminosity As Integer, fillSaturation As Integer

Public CapType As ttCapType
Public JoinType As ttJoinType
Public cFillTransparency As Double
Public cPenTransparency As Double
Public penSize As Double
Dim cPenColor As ttColors
Public DrawingMode As ttDrawingMode
Public FillType As ttFillType
Public LineStyle As ttLineStyle
Dim cFillColor As ttColors
Public FontSize As Long
Public FontName As String
Public FontColor As ttColors
Public FontStyle As ttFontStyle
Public FontWeight As ttFontWeight

Private Sub RGB2HSL(RGBColor As Long, Hue As Integer, Saturation As Integer, Luminosity As Integer)
  ColorRGBToHLS RGBColor, Hue, Luminosity, Saturation
  Hue = (360 * (Hue / 239)) Mod 360
End Sub

Private Function HSL2RGB(Hue As Integer, Saturation As Integer, Luminosity As Integer)
  Dim h As Long
  h = Hue / 360 * 239
  HSL2RGB = ColorHLSToRGB(h, Luminosity, Saturation)
End Function


Public Function ImageLost() As Boolean
  Dim l As Double
  ImageLost = False
  If TurtleImage Is Nothing Then
    ImageLost = True
    Exit Function
  End If
  On Error GoTo SetImageLost:

  l = TurtleImage.left
  Exit Function
SetImageLost:
    ImageLost = True
End Function

Public Function InitTurtle(ByVal CanvasName As String, ByVal TurtleName As String) As Integer
  Set TurtleImage = ActiveSheet.ChartObjects(CanvasName).Chart.Shapes(TurtleName)
  TurtleImage.ZOrder msoBringToFront

  nAngle = INITIAL_ANGLE
  Set Canvas = ActiveSheet.ChartObjects(CanvasName).Chart
  Set CanvasShape = ActiveSheet.Shapes(CanvasName)
  penSize = 1
  penColor = 0
  IsPenDown = True
  FontSize = 11
  FontName = "Arial"
  FontColor = ttblack
  FontStyle = ttFSNormal
  FontWeight = ttFWNormal
  fillColor = ttinvisible
  RefreshLines = 10
  DrawingMode = ttScreenRefresh
  IsCurvedPoly = False
  Me.Center
End Function

Private Sub createSAoPIfNotExists(ByVal X As Double, ByVal Y As Double, Optional nodes As Long = 1)
    If PolyNodes = 0 Then
      BufLimit = 1000
      ReDim SAoP(1 To BufLimit)
      SAoP(1) = Array(X, Y)
      PolyNodes = 1
    End If
    If PolyNodes + nodes > BufLimit Then
      BufLimit = BufLimit + 1000
      ReDim Preserve SAoP(1 To BufLimit)
    End If
End Sub

Private Sub getStepsAngle(ByVal X As Double, ByVal Y As Double, ByRef steps As Double, ByRef angle As Double)

  steps = Sqr((X - PenX) ^ 2 + (Y - PenY) ^ 2)
  If Y - PenY <> 0 Then
    angle = Atn((X - PenX) / -(Y - PenY)) * 180 / [pi()]
  Else
    If X < PenX Then
      angle = -90
    Else
      angle = 90
    End If
  End If
  If Y > PenY Then
    angle = angle + 180
  End If
  If angle = 0 And Y > PenY Then
    angle = 180
  End If
End Sub



Public Sub MoveXY(ByVal X As Double, Y As Double)
  Dim original_angle As Double, steps As Double
  
  original_angle = nAngle
  getStepsAngle X, Y, steps, nAngle
  Move steps
  PenX = X
  PenY = Y
  nAngle = original_angle
End Sub


Public Sub Move(ByVal steps As Double)
'
' Add a Line to the drawing
  

  Dim newX As Double, newY As Double, line As shape, i As Long
  
  If IsCurvedPoly Then
    MoveCurved steps, 0, ttLine
  Else

    newX = PenX + Math.Cos((nAngle - INITIAL_ANGLE) * PI / 180) * steps
    newY = PenY + Math.Sin((nAngle - INITIAL_ANGLE) * PI / 180) * steps
    
    NumLines = NumLines + 1
    If IsPenDown Then
      IsCurvedPoly = False
      createSAoPIfNotExists PenX, PenY
      PolyNodes = PolyNodes + 1
      SAoP(PolyNodes) = Array(newX, newY)
         
      If DrawingMode = ttScreenRefresh Then
        RecreatePoly
      End If
    End If
    PenX = newX
    PenY = newY
    RepositionTurtle
  End If
End Sub



Private Sub RecreatePoly()
  Dim i As Long, previous_size As Long
  

  If Not (tempPoly Is Nothing) Then
    'it is much faster to recreate the polyline
    'than to adding nodes with nodes.insert
    tempPoly.Delete
    Set tempPoly = Nothing
  End If
  
  'store previous safearray size
  previous_size = UBound(SAoP)
  ReDim Preserve SAoP(1 To PolyNodes)
  If IsCurvedPoly Then
    Set tempPoly = Canvas.Shapes.AddCurve(SAoP)
  Else
    If UBound(SAoP) > 2 Then
      Set tempPoly = Canvas.Shapes.AddPolyline(SAoP)
    Else
      Set tempPoly = Canvas.Shapes.AddLine(SAoP(1)(0), SAoP(1)(1), SAoP(2)(0), SAoP(2)(1))
    End If
  End If
  FormatShape tempPoly, False
  If NumLines > RefreshLines Or PolyNodes < RefreshLines And PolyNodes > 2 Then
    DoEvents
    NumLines = 0
  End If
  'restore previous safearray size
  ReDim Preserve SAoP(1 To previous_size)
End Sub




Public Sub Reset()
  Me.Clear
  Me.PointInDirection 90
  Me.Center
  Me.Show
  Me.PenDown
  Me.DrawingMode = ttScreenRefresh
  Me.penSize = 1
  Me.fillColor = ttinvisible
  Me.penColor = ttblack
  Me.CanvasColor = ttwhite
  Me.PenTransparency = 0
  Me.FillTransparency = 0
  Me.JoinType = ttMiter
  Me.CapType = ttroundcap
  Set tempPoly = Nothing
  Set PolyPoints = Nothing
  PolyNodes = 0
End Sub

Public Sub Center()
  If IsPenDown Then
    Me.PenUp
    Me.PenDown
  End If
  PenX = Canvas.ChartArea.width / 2
  PenY = Canvas.ChartArea.height / 2
  RepositionTurtle
End Sub

Public Sub PenDown()
  If Not IsPenDown And PolyNodes > 0 Then
      If Not (tempPoly Is Nothing) Then
        tempPoly.Delete
      End If
      Set tempPoly = Nothing
      PolyNodes = 0
  End If
  IsPenDown = True
End Sub

Public Sub PenHueShift(ByVal Shift As Long)
  penHue = penHue + Shift
  cPenColor = HSL2RGB(penHue, penSaturation, penLuminosity)

End Sub
Public Sub FillHueShift(ByVal Shift As Long)
  fillHue = fillHue + Shift
  cFillColor = HSL2RGB(fillHue, fillSaturation, fillLuminosity)

End Sub


Public Function PenUp() As ShapeRange
  Dim line As shape
  If IsPenDown Then
    If PolyNodes > 0 Then
      If Not (tempPoly Is Nothing) Then
        tempPoly.Delete
      End If
      Set tempPoly = Nothing
      ' if first point and last one look like the same up to
      ' the 7th decimal, make them the same
      If Round(SAoP(PolyNodes)(0), 7) = Round(SAoP(1)(0), 7) Then
        SAoP(PolyNodes)(0) = SAoP(1)(0)
      End If
      If Round(SAoP(PolyNodes)(1), 7) = Round(SAoP(1)(1), 7) Then
        SAoP(PolyNodes)(1) = SAoP(1)(1)
      End If

      RecreatePoly
      FormatShape tempPoly
      PolyNodes = 0
      Set PenUp = Canvas.Shapes.Range(Array(tempPoly.Name))
    End If
    IsPenDown = False
    Set tempPoly = Nothing
    PolyNodes = 0
    IsCurvedPoly = False
  End If
  'DoEvents
End Function




Public Sub Clear()
   Dim shpFrame As shape
   Dim intIndex As Integer
   If Not (Canvas Is Nothing) Then
      For Each shpFrame In Canvas.Shapes
        If LCase(shpFrame.Name) <> "turtle" Then
          shpFrame.Delete
        End If
      Next
    End If
  Set tempPoly = Nothing
  PolyNodes = 0
End Sub

Private Sub FormatShape(shp As shape, Optional bFill As Boolean = True)
  With shp
    
    If CapType <> ttroundcap Or JoinType <> ttMiter Then
      Dim reference As Variant, refShape As shape
      reference = Array("round_round", "flat_round", "square_round", _
                        "round_bevel", "flat_bevel", "square_bevel", _
                        "round_miter", "flat_miter", "square_miter")
      
      Set refShape = Worksheets("Canvas").Shapes(reference(CapType + JoinType * 3))
      refShape.PickUp
      .Apply
 
    End If
    .line.ForeColor.RGB = penColor
    .line.Transparency = cPenTransparency
    If fillColor <> ttinvisible And bFill And .nodes.Count > 2 Then
      .Fill.ForeColor.RGB = fillColor
      .Fill.Visible = msoTrue
      .Fill.Solid
      .Fill.Transparency = cFillTransparency
    Else
      If .Fill.Visible <> msoFalse Then
        .Fill.Visible = msoFalse
      End If
    End If
    
    .line.Weight = penSize
    If penColor = ttinvisible Then
      .line.Visible = msoFalse
    End If
    If TurtleImage.Visible Then
      TurtleImage.ZOrder msoBringToFront '.ZOrder msoSendToBack
    End If
'    .Placement = xlFreeFloating
    .LockAspectRatio = msoCTrue
  End With
End Sub


Public Function Group(Optional numShapes As Long, Optional ungroupFirst As Boolean = True) As ShapeRange
'
' Group all shapes
'
  Dim shpFrame As shape
  Dim intIndex As Integer, i As Long
    
  If ungroupFirst Then
    For Each shpFrame In Canvas.Shapes
      If shpFrame.Type = 6 Then
        shpFrame.Ungroup
      End If
    Next
  End If

  intIndex = 0
  ReDim vntitems(Canvas.Shapes.Count - 1) As Variant
  For i = 1 To Canvas.Shapes.Count
      Set shpFrame = Canvas.Shapes(i)
      If shpFrame.Type <> msoPicture Then
        vntitems(intIndex) = i
        
        intIndex = intIndex + 1
      End If
  Next
  If intIndex > 1 Then
    If intIndex < Canvas.Shapes.Count Then
      ReDim Preserve vntitems(intIndex - 1)
    End If
    If numShapes <> 0 And numShapes < intIndex Then
      For i = intIndex - numShapes To intIndex - 1
        vntitems(i - intIndex + numShapes) = vntitems(i)
      Next i
      ReDim Preserve vntitems(numShapes - 1)
    End If
    If UBound(vntitems) > 0 Then
        Set shpFrame = Canvas.Shapes.Range(vntitems).Group()
        Set Group = Canvas.Shapes.Range(Array(shpFrame.Name))
 
    Else
      Set Group = Canvas.Shapes.Range(vntitems(0))
    End If
  ElseIf intIndex = 1 Then
    Set Group = Canvas.Shapes.Range(vntitems(0))

  End If
End Function

Public Sub Ungroup()
  Dim shpFrame As shape
  For Each shpFrame In Canvas.Shapes
    If shpFrame.Type = msoGroup Then
      shpFrame.Ungroup
    End If
  Next
End Sub

Public Sub TurnRight(ByVal angle As Double)
  Dim dec As Double
  nAngle = nAngle + angle
  TurtleImage.IncrementRotation angle

End Sub

Public Sub TurnLeft(ByVal angle As Double)
  TurtleImage.IncrementRotation -angle
  nAngle = nAngle - angle

End Sub

Public Sub PointInDirection(ByVal angle As Double)
  TurtleImage.rotation = angle - INITIAL_ANGLE
  nAngle = angle
End Sub


Public Sub Show()
  TurtleImage.Visible = msoTrue
  RepositionTurtle
End Sub

Public Sub Hide()
  'TurtleImage.SoftEdge.Radius = 100
  TurtleImage.Visible = msoFalse
End Sub

Public Sub GoToXY(ByVal X As Double, ByVal Y As Double)
  If IsPenDown Then
    Me.PenUp
    Me.PenDown
  End If
  PenX = X
  PenY = Y
  RepositionTurtle
End Sub

Public Property Get CanvasWidth() As Double
    CanvasWidth = Canvas.ChartArea.width
End Property

Private Property Get width() As Double
    width = TurtleImage.width
End Property
Private Property Get height() As Double
    height = TurtleImage.height
End Property

Public Property Get CanvasHeight() As Double
    CanvasHeight = Canvas.ChartArea.height
End Property

Public Property Get X() As Double
  X = PenX
End Property
Public Property Let X(ByVal value As Double)
 PenX = X
 RepositionTurtle
End Property

Public Property Get Y() As Double
   Y = PenY
End Property
Public Property Let Y(ByVal value As Double)
  PenY = Y
  RepositionTurtle
End Property
 
Public Sub Wait(milliseconds As Long)
  Pause milliseconds / 1000
End Sub

Public Property Let CanvasColor(value As ttColors)
  If value = ttinvisible Then
    CanvasShape.Fill.Visible = msoFalse
  Else
    CanvasShape.Fill.ForeColor.RGB = value
  End If
End Property
 
Public Property Get CanvasColor() As ttColors
 CanvasColor = CanvasShape.Fill.ForeColor.RGB
End Property
Public Property Let penColor(value As ttColors)
  RGB2HSL value, penHue, penSaturation, penLuminosity
  cPenColor = value
End Property
 
Public Property Get penColor() As ttColors
  penColor = cPenColor
End Property
 
Public Property Let fillColor(value As ttColors)
  RGB2HSL value, fillHue, fillSaturation, fillLuminosity
  cFillColor = value
End Property
 
Public Property Get fillColor() As ttColors
  fillColor = cFillColor
End Property
Function Pause(ByVal NumberOfSeconds As Variant)
    On Error GoTo Error_GoTo

    Dim PauseTime As Variant
    Dim start As Variant
    Dim Elapsed As Variant

    PauseTime = NumberOfSeconds
    start = Timer
    Elapsed = 0
    Do While Timer < start + PauseTime
        Elapsed = Elapsed + 1
        If Timer = 0 Then
            ' Crossing midnight
            PauseTime = PauseTime - Elapsed
            start = 0
            Elapsed = 0
        End If
        DoEvents
    Loop

Exit_GoTo:
    On Error GoTo 0
    Exit Function
Error_GoTo:
    Debug.Print Err.Number, Err.Description, Erl
    GoTo Exit_GoTo
End Function

Public Sub Point()
  If PolyPoints Is Nothing Then
    Set PolyPoints = Canvas.Shapes.BuildFreeform(msoEditingCorner, PenX, PenY)
    InitialPenY = PenY
    InitialPenX = PenX
  Else
    PolyPoints.AddNodes msoSegmentLine, msoEditingCorner, PenX, PenY
  End If
    
End Sub

Public Sub ClosePoints(Optional sendtoBack As Long = 0)
  Dim line As shape, i As Long
  If Not (PolyPoints Is Nothing) Then
    PolyPoints.AddNodes msoSegmentLine, msoEditingCorner, InitialPenX, InitialPenY 'close the poly
    Set line = PolyPoints.ConvertToShape
    Set PolyPoints = Nothing
    FormatShape line
    If sendtoBack Then
      For i = 1 To sendtoBack
        line.ZOrder msoSendBackward
      Next i
    End If
    
  End If
End Sub

Public Sub Ellipse(ByVal DiameterAcross As Double, Optional DiameterFrontBack)
  Dim shp As shape
  If IsPenDown Then
    If IsMissing(DiameterFrontBack) Then DiameterFrontBack = DiameterAcross
    Set shp = Canvas.Shapes.AddShape(msoShapeOval, Me.X - DiameterFrontBack / 2, Me.Y - DiameterAcross / 2, DiameterFrontBack, DiameterAcross)
    shp.rotation = TurtleImage.rotation
    FormatShape shp
    If DrawingMode = ttScreenRefresh Then
      NumLines = NumLines + 1
      If NumLines > RefreshLines Then
        DoEvents
        NumLines = 0
      End If
    End If
  End If
End Sub

Public Sub Arc(ByVal DiameterAcross As Double, ByVal DiameterFrontBack As Double, ByVal StartAngle As Double, ByVal EndAngle As Double, ByVal ArcType As ttArcType)
  Dim shp As shape
  If IsPenDown Then
    If ArcType = TTARC Then
      Set shp = Canvas.Shapes.AddShape(msoShapeArc, Me.X, Me.Y - DiameterAcross / 2, DiameterFrontBack / 2, DiameterAcross / 2)
    ElseIf ArcType = ttsector Then
      Set shp = Canvas.Shapes.AddShape(msoShapePie, Me.X - DiameterFrontBack / 2, Me.Y - DiameterAcross / 2, DiameterFrontBack, DiameterAcross)
    ElseIf ArcType = ttSegment Then
      Set shp = Canvas.Shapes.AddShape(msoShapeChord, Me.X - DiameterFrontBack / 2, Me.Y - DiameterAcross / 2, DiameterFrontBack, DiameterAcross)
    End If
    shp.Adjustments.Item(1) = StartAngle - INITIAL_ANGLE
    shp.Adjustments.Item(2) = EndAngle - INITIAL_ANGLE
    shp.rotation = TurtleImage.rotation
    FormatShape shp
  End If
End Sub


Public Function ColorPicker(Optional copy As Boolean = False)
  Dim document As HTMLDocument, color As Long
  Dim Red As Long, Green As Long, Blue As Long
  
  Red = Me.fillColor Mod 256
  Green = (Me.fillColor \ 256) Mod 256
  Blue = Me.fillColor \ 65536
  If Red = -1 Then
    Red = 255
    Green = 255
    Blue = 255
  End If
  If Application.Dialogs(xlDialogEditColor).Show(1, Red, Green, Blue) = True Then
    color = ActiveWorkbook.Colors(1)
    If copy Then
      Set document = CreateObject("HtmlFile")
      document.parentWindow.clipboardData.SetData "text", Str(color)
    End If
    ColorPicker = color
  End If
End Function

Public Function ColorDropper(Optional copy As Boolean = False)
  Dim document As HTMLDocument, color As Long
  Dim Red As Long, Green As Long, Blue As Long
  
  Red = Me.fillColor Mod 256
  Green = (Me.fillColor \ 256) Mod 256
  Blue = Me.fillColor \ 65536
  If Red = -1 Then
    Red = 255
    Green = 255
    Blue = 255
  End If
  If Me.fillColor <> -1 Then
    frmEyeDropper.imgColor.BackColor = Me.fillColor
  End If
  frmEyeDropper.Show vbModal
  ColorDropper = frmEyeDropper.imgColor.BackColor
  If copy Then
      Set document = CreateObject("HtmlFile")
      document.parentWindow.clipboardData.SetData "text", Str(ColorDropper)
  End If
End Function

Public Sub WriteText(txt As String)
  Dim shp As shape
 
  Set shp = Canvas.Shapes.AddTextbox(msoTextOrientationHorizontal, Me.X, Me.Y, 1, 1)
  shp.TextFrame2.TextRange.Characters.Text = txt
  shp.TextFrame2.TextRange.Font.Name = FontName
  shp.TextFrame2.TextRange.Font.Size = FontSize
  If FontStyle = ttfsItalic Then
    shp.TextFrame2.TextRange.Font.Italic = msoTrue
  Else
    shp.TextFrame2.TextRange.Font.Italic = msoFalse
  End If
  If FontWeight = ttFWBold Then
    shp.TextFrame2.TextRange.Font.Bold = msoCTrue
  Else
    shp.TextFrame2.TextRange.Font.Bold = msoFalse
  End If
  shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = FontColor
  shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText
  shp.TextFrame2.WordWrap = msoFalse
  shp.left = Me.X - shp.width / 2
  shp.top = Me.Y - shp.height / 2
  shp.rotation = TurtleImage.rotation


End Sub


Sub SaveCanvas(fileName As String, ImageFormat As ttImageFormat)
  Dim tmpVisible As Boolean, shpFrame As shape
  
  
  If Canvas.Parent.Name = "Canvas2" Then
    
    tmpVisible = TurtleImage.Visible
    TurtleImage.Visible = False
    If ImageFormat = ttPNG Then
      Canvas.Export fileName, "PNG"
    ElseIf ImageFormat = ttGIF Then
      Canvas.Export fileName, "GIF"
    ElseIf ImageFormat = ttJPG Then
      Canvas.Export fileName, "JPG"
    ElseIf ImageFormat = ttPDF Then
      Canvas.ExportAsFixedFormat xlTypePDF, fileName
    ElseIf ImageFormat = ttEMF Then
      Canvas.ChartArea.Select
      Excel.Selection.copy
      fnSaveAsEMF fileName
    End If
  
    
    TurtleImage.Visible = tmpVisible
  Else
    Dim canvas_copy As ChartObject
    Dim sht As Worksheet
    Dim tt As New clsTurtle, o As Object
    Dim s As ShapeRange, xywh As Variant, l As Double, t As Double
    
    Set sht = Worksheets("Canvas")
    'duplicate canvas
    sht.Shapes("Canvas").Duplicate
    Set canvas_copy = sht.ChartObjects(sht.ChartObjects.Count)
    canvas_copy.left = 0
    canvas_copy.Name = "Canvas2"
    tt.InitTurtle "Canvas2", "Turtle"
    Set s = tt.Group
    xywh = get_xywh(s)
    l = s.left - xywh(0)
    t = s.top - xywh(1)
    s.Select
    Excel.Selection.Cut
    tt.CanvasColor = ttinvisible
    canvas_copy.Chart.ChartArea.width = xywh(2) + 2
    'canvas_copy.height = h
    canvas_copy.Chart.ChartArea.height = xywh(3) + 2
    
    canvas_copy.Chart.Paste
    canvas_copy.Chart.Shapes(canvas_copy.Chart.Shapes.Count).left = l
    canvas_copy.Chart.Shapes(canvas_copy.Chart.Shapes.Count).top = t
    If ImageFormat = ttPDF Then
      Canvas.Parent.Visible = xlSheetHidden
      For Each o In sht.OLEObjects
        o.Visible = xlSheetHidden
      Next o
      'resize to keep proportions in pdf
      canvas_copy.Chart.ChartArea.height = (xywh(3)) * 1.0913907 '  1.106254
    End If

    tt.SaveCanvas fileName, ImageFormat
    canvas_copy.Select
    Excel.Selection.Delete
    Dim document As New HTMLDocument
    document.parentWindow.clipboardData.clearData
    If ImageFormat = ttPDF Then
      Canvas.Parent.Visible = xlSheetVisible
      For Each o In sht.OLEObjects
        o.Visible = xlSheetVisible
      Next o
    End If
  
  
  End If
End Sub


' see http://stackoverflow.com/questions/1791369/excel-export-chart-to-wmf-or-emf
Private Function fnSaveAsEMF(strFileName As String) As Boolean
Const CF_ENHMETAFILE As Long = 14

Dim ReturnValue As Long

    OpenClipboard 0

    ReturnValue = CopyEnhMetaFileA(GetClipboardData(CF_ENHMETAFILE), strFileName)

    EmptyClipboard

    CloseClipboard

    '// Release resources to it eg You can now delete it if required
    '// or write over it. This is a MUST
    DeleteEnhMetaFile ReturnValue

    fnSaveAsEMF = (ReturnValue <> 0)

End Function

Private Sub arcCircle(ByVal steps As Double, ByVal radius As Double)
      Dim xe As Double, ye As Double, positive As Boolean
      Dim xc As Double, yc As Double
      
      xe = PenX + Math.Cos((nAngle - INITIAL_ANGLE) * [pi()] / 180) * steps          ' x-end
      ye = PenY + Math.Sin((nAngle - INITIAL_ANGLE) * [pi()] / 180) * steps
     
      If Abs(radius) < Abs(steps) / 2 Then
        Err.Raise vbObjectError + 513, "Turtle Graphics", _
           "Radius has to be greater than half the steps." + vbNewLine + _
           "Radius: " + Str(radius) + " Steps: " + Str(steps)
      ElseIf radius > 0 Then
        positive = True
      Else
        positive = False
        radius = -radius
      End If
      getCenterOfCircle PenX, PenY, xe, ye, radius, positive, xc, yc
      
      'from https://stackoverflow.com/questions/734076/how-to-best-approximate-a-geometrical-arc-with-a-bezier-curve
      Dim ax As Double, ay As Double, bx As Double, by As Double
      Dim q1 As Double, q2 As Double, k2 As Double, x2 As Double, y2 As Double, x3 As Double, y3 As Double
      ax = PenX - xc
      ay = PenY - yc
      bx = xe - xc
      by = ye - yc
      q1 = ax * ax + ay * ay
      q2 = q1 + ax * bx + ay * by
      k2 = (4 / 3) * (Sqr(2 * q1 * q2) - q2) / (ax * by - ay * bx)
      
      x2 = xc + ax - k2 * ay
      y2 = yc + ay + k2 * ax
      x3 = xc + bx + k2 * by
      y3 = yc + by - k2 * bx
      If IsPenDown Then
        IsCurvedPoly = True
        createSAoPIfNotExists PenX, PenY, 3
        
        PolyNodes = PolyNodes + 1
        SAoP(PolyNodes) = Array(x2, y2)
        
        PolyNodes = PolyNodes + 1
        SAoP(PolyNodes) = Array(x3, y3)
        
        PolyNodes = PolyNodes + 1
        SAoP(PolyNodes) = Array(xe, ye)
       
        If DrawingMode = ttScreenRefresh Then
          RecreatePoly
        End If
      End If
      PenX = xe
      PenY = ye
      RepositionTurtle
End Sub

Sub MoveXYCurved(ByVal X As Double, ByVal Y As Double, ByVal radius As Double, _
    Optional curveType As ttCurveType = ttHalfEllipse)
  Dim original_angle As Double, steps As Double
  
  original_angle = nAngle
  getStepsAngle X, Y, steps, nAngle
  MoveCurved steps, radius, curveType
  PenX = X
  PenY = Y
  nAngle = original_angle
    
End Sub
Sub MoveCurved(ByVal steps As Double, ByVal radius As Double, _
    Optional curveType As ttCurveType = ttHalfEllipse)
  Dim EllipseAngle As Double, w As Double, StepsForward As Double
  Dim angle1 As Double, angle2 As Double, originalAngle As Double
  
   
  If Not IsCurvedPoly And PolyNodes > 0 And IsPenDown Then
    Me.PenUp
    Me.PenDown
  End If
  originalAngle = nAngle
  'Treat steps < 0 as a rotation
  If steps < 0 Then
    nAngle = nAngle + 180
    steps = -steps
  End If
  If curveType = ttarccircle And Abs(radius) = Abs(steps / 2) Then
    curveType = ttHalfEllipse
  End If
  
  If curveType = ttHalfEllipse Then
    w = steps / 2
    MoveTangent w, radius, 0, 0
    If radius > 0 Then
      nAngle = nAngle + 90
    Else
      nAngle = nAngle - 90
      w = -w
      radius = -radius
    End If
    
    MoveTangent radius, w, 0, 0

    If w > 0 Then
      nAngle = nAngle - 90
    Else
      nAngle = nAngle + 90
    End If

  ElseIf curveType = ttPetalfd Or curveType = ttPetalbk Then
      If curveType = ttPetalfd Then
        w = steps * 0.6
        angle1 = 45
        angle2 = 0
      Else
        w = steps * 0.4
        angle1 = 0
        angle2 = -45
      End If
      
      MoveTangent w, radius, angle1, 0
    
      w = steps - w
      
      If radius > 0 Then
        nAngle = nAngle + 90
      Else
        nAngle = nAngle - 90
      End If

      MoveTangent radius, w, 0, angle2

      If radius > 0 Then
        nAngle = nAngle - 90
      Else
        nAngle = nAngle + 90
      End If
  ElseIf curveType = ttQuarterEllipse Then
      StepsForward = Sqr(steps ^ 2 - radius ^ 2)
      EllipseAngle = Math.Atn(radius / StepsForward) / [pi()] * 180
      nAngle = nAngle + EllipseAngle
      MoveTangent StepsForward, radius, 0, 0
  ElseIf curveType = ttarccircle Then
      arcCircle steps, radius
  ElseIf curveType = ttCusp Then
      nAngle = nAngle - 90
      MoveTangent radius, -steps / 2, 90, 0, 0.55, 0.55
      MoveTangent -radius, steps / 2, 90, 0, 0.55, 0.55
      nAngle = nAngle + 90
  ElseIf curveType = ttLine Then
      MoveTangent steps, 0, 90, 0, 0, 0
  End If
      
      
  
  nAngle = originalAngle
End Sub


Public Function getSideLength(ByVal radius As Double, ByVal sides As Double)
  getSideLength = radius * (2 * Sin([pi()] / sides))
End Function

Public Function getRadiusLength(ByVal sideLength As Double, ByVal sides As Long)
  getRadiusLength = sideLength / (2 * Sin([pi()] / sides))
End Function


Private Sub getStepsForwardAndLeft(ByVal angle, ByVal radius, StepsForward As Double, StepsLeft As Double)
  StepsForward = radius * Cos(angle / 180 * [pi()])
  StepsLeft = radius * Sin(angle / 180 * [pi()])
End Sub


Private Sub RepositionTurtle()
  If TurtleImage.Visible Then
    TurtleImage.left = PenX - TurtleImage.width / 2
    TurtleImage.top = PenY - TurtleImage.height / 2
    TurtleImage.ZOrder msoBringToFront
  End If
End Sub


Public Sub MoveTangent(ByVal StepsForward As Double, ByVal StepsLeft As Double, _
        ByVal FirstAngleOffset As Double, ByVal SecondAngleOffset As Double, _
        Optional factor1, Optional factor2)
  Dim ellipse_angle As Double, kappa As Double, chord As Double
  Dim X As Double, Y As Double, xe As Double, ye As Double, ox As Double, oy As Double
  Dim cx1 As Double, cy1 As Double, cx2 As Double, cy2 As Double
  Dim originalAngle As Double, restoreAngle As Boolean
  
  restoreAngle = False
  If StepsForward < 0 Then
    originalAngle = nAngle
    nAngle = nAngle + 180
    StepsForward = -StepsForward
    restoreAngle = True
  End If
  
  kappa = 0.5522848
  
  ellipse_angle = Math.Atn(StepsLeft / StepsForward) / [pi()] * 180
  chord = Sqr(StepsForward ^ 2 + StepsLeft ^ 2)
  'turn in direction of the ellipse co-vertex
  nAngle = nAngle - ellipse_angle
  X = Me.X
  Y = Me.Y
  
  ' control points use by default
  ' the proportion of the ellipse
  ' based on the factor kappa of the opposite movement
  ' if factor informed, it will use the proportion of the
  ' forward direction
  
  'control point offset horizontal
  If IsMissing(factor1) Then
    ox = StepsForward * kappa
  Else
    ox = StepsForward * factor1
  End If
  
  'control point offset vertical
  If IsMissing(factor2) Then
    oy = StepsLeft * kappa
  Else
    oy = StepsForward * factor2
  End If
  xe = X + Math.Cos((nAngle - INITIAL_ANGLE) * [pi()] / 180) * chord          ' x-end
  ye = Y + Math.Sin((nAngle - INITIAL_ANGLE) * [pi()] / 180) * chord

  cx1 = X - Math.Sin([pi()] / 180 * (90 - nAngle - ellipse_angle - FirstAngleOffset)) * oy
  cy1 = Y - Math.Cos([pi()] / 180 * (90 - nAngle - ellipse_angle - FirstAngleOffset)) * oy
  
  cx2 = xe - Math.Cos([pi()] / 180 * (90 - nAngle - ellipse_angle - SecondAngleOffset)) * ox
  cy2 = ye + Math.Sin([pi()] / 180 * (90 - nAngle - ellipse_angle - SecondAngleOffset)) * ox
  
    NumLines = NumLines + 1
  If IsPenDown Then
    IsCurvedPoly = True
    createSAoPIfNotExists X, Y, 3
    
    PolyNodes = PolyNodes + 1
    SAoP(PolyNodes) = Array(cx1, cy1)
    
    PolyNodes = PolyNodes + 1
    SAoP(PolyNodes) = Array(cx2, cy2)
   
    PolyNodes = PolyNodes + 1
    SAoP(PolyNodes) = Array(xe, ye)
    
    If DrawingMode = ttScreenRefresh Then
      RecreatePoly
    End If

  End If
  PenX = xe
  PenY = ye
  RepositionTurtle
  
  ' restore original direction
  nAngle = nAngle + ellipse_angle
  If restoreAngle Then
     nAngle = originalAngle
  End If
End Sub

Public Sub MoveXYBezier(ByVal X As Double, ByVal Y As Double, ByVal angle1 As Double, _
        ByVal distance1 As Double, ByVal angle2 As Double, _
        ByVal distance2 As Double)

  Dim original_angle As Double, steps As Double
  
  original_angle = nAngle
  getStepsAngle X, Y, steps, nAngle
  MoveBezier steps, angle1, distance1, angle2, distance2
  PenX = X
  PenY = Y
  nAngle = original_angle


End Sub

Public Sub MoveBezier(ByVal steps As Double, angle1 As Double, _
        ByVal distance1 As Double, ByVal angle2 As Double, _
        ByVal distance2 As Double)
  
  Dim X As Double, Y As Double, xe As Double, ye As Double, ox1 As Double, oy1 As Double
  Dim ox2 As Double, oy2 As Double
  Dim cx1 As Double, cy1 As Double, cx2 As Double, cy2 As Double
  Dim originalAngle As Double, restoreAngle As Boolean
  
  If steps = 0 Then
    Exit Sub
  End If
  restoreAngle = False
  If steps < 0 Then
    originalAngle = nAngle
    nAngle = nAngle + 180
    steps = -steps
    restoreAngle = True
  End If
  
  X = Me.X
  Y = Me.Y
  
  ' control points
  
  'control points offset horizontal
  ox1 = distance1 * Cos((nAngle - INITIAL_ANGLE + angle1) * [pi()] / 180)
  ox2 = distance2 * Cos((nAngle - INITIAL_ANGLE + angle2) * [pi()] / 180)
  
  'control points offset vertical
  oy1 = distance1 * Sin((nAngle - INITIAL_ANGLE + angle1) * [pi()] / 180)
  oy2 = distance2 * Sin((nAngle - INITIAL_ANGLE + angle2) * [pi()] / 180)
  
  
  xe = X + Math.Cos((nAngle - INITIAL_ANGLE) * [pi()] / 180) * steps          ' x-end
  ye = Y + Math.Sin((nAngle - INITIAL_ANGLE) * [pi()] / 180) * steps

  cx1 = X + ox1
  cy1 = Y + oy1
  
  cx2 = xe + ox2
  cy2 = ye + oy2
  
    NumLines = NumLines + 1
  If IsPenDown Then
    IsCurvedPoly = True
    createSAoPIfNotExists X, Y, 3
    
    PolyNodes = PolyNodes + 1
    SAoP(PolyNodes) = Array(cx1, cy1)
    
    PolyNodes = PolyNodes + 1
    SAoP(PolyNodes) = Array(cx2, cy2)
   
    PolyNodes = PolyNodes + 1
    SAoP(PolyNodes) = Array(xe, ye)
    
    If DrawingMode = ttScreenRefresh Then
      RecreatePoly
    End If

  End If
  PenX = xe
  PenY = ye
  RepositionTurtle
  
  ' restore original direction
  If restoreAngle Then
     nAngle = originalAngle
  End If
End Sub


' From https://rosettacode.org/wiki/Circles_of_given_radius_through_two_points#VBA
Private Sub getCenterOfCircle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, _
      ByVal y2 As Double, ByVal R As Double, ByVal positive As Boolean, xc As Double, yc As Double)
  Dim xd As Double, yd As Double, s2 As Double, sep As Double
  Dim xh As Double, yh As Double, md As Double, xs As Double, ys As Double

        xd = x2 - x1
        yd = y1 - y2
        s2 = xd * xd + yd * yd
        sep = Sqr(s2)
        xh = (x1 + x2) / 2
        yh = (y1 + y2) / 2
        Dim txt As String
        If sep = 0 Then
            Debug.Print "same points/" & IIf(R = 0, "radius is zero", "infinite solutions")
        Else
            If sep = 2 * R Then
                Debug.Print "opposite ends of diameter with centre " & xh & ", " & yh & "."
            Else
                If sep > 2 * R Then
                    Debug.Print "too far apart " & sep & " > " & 2 * R
                Else
                    md = Sqr(R * R - s2 / 4)
                    xs = md * xd / sep
                    ys = md * yd / sep
                    If positive Then
                      xc = xh + ys
                      yc = yh + xs
                    Else
                      xc = xh - ys
                      yc = yh - xs
                    End If
                End If
            End If
        End If
        
End Sub

Public Property Get FillTransparency() As Long
  FillTransparency = cFillTransparency * 100
End Property

Public Property Let FillTransparency(ByVal NewValue As Long)
 cFillTransparency = NewValue / 100
End Property

Public Property Get PenTransparency() As Long
  PenTransparency = cPenTransparency * 100
End Property

Public Property Let PenTransparency(ByVal NewValue As Long)
 cPenTransparency = NewValue / 100
End Property

Private Function get_xywh(sr As ShapeRange) As Variant
   ' calculate new top and left for printing
  Dim shp As shape, xmin As Double, ymin As Double, vert1 As Variant, vert2 As Variant
  Dim rotation As Double, sin_ang As Double, xmax As Double, ymax As Double
  Dim cos_ang As Double, i As Long, cx As Double, cy As Double
  xmin = sr.left
  ymin = sr.top
  xmax = sr.left + sr.width
  ymax = sr.top + sr.height
  If sr.Type <> msoGroup Then
    get_xywh = Array(xmin, ymin, xmax - xmin, ymax - ymin)
    Exit Function
  End If
  For Each shp In sr.GroupItems
    rotation = shp.rotation / 180 * [pi()]
    cx = shp.left + shp.width / 2
    cy = -shp.top - shp.height / 2
    
    
    vert1 = Array(Array(shp.left - cx, -shp.top - cy), _
                  Array(shp.left + shp.width - cx, -shp.top - cy), _
                  Array(shp.left - cx, -shp.top - shp.height - cy), _
                  Array(shp.left + shp.width - cx, -shp.top - shp.height - cy))
    vert2 = vert1
     sin_ang = Sin(rotation)
     cos_ang = Cos(rotation)
     For i = 0 To 3
      vert2(i)(0) = vert1(i)(0) * cos_ang - vert1(i)(1) * sin_ang + cx
      vert2(i)(1) = vert1(i)(0) * sin_ang + vert1(i)(1) * cos_ang + cy
     Next i
     For i = 0 To 3
      If vert2(i)(0) < xmin Then xmin = vert2(i)(0)
      If -vert2(i)(1) < ymin Then ymin = -vert2(i)(1)
      If vert2(i)(0) > xmax Then xmax = vert2(i)(0)
      If -vert2(i)(1) > ymax Then ymax = -vert2(i)(1)
     Next i
  Next shp
  get_xywh = Array(xmin, ymin, xmax - xmin, ymax - ymin)
End Function



