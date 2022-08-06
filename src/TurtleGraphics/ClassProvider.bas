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
