VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private Sub cmdColorPicker_Click()
  turtle.ColorPicker True
End Sub
Private Sub cmdEdit_Click()
  Application.Goto "YourProgramHere"
End Sub

Private Sub cmdReset_Click()
  turtle.Reset
End Sub



Private Sub cmdExport_Click()
  Dim FileName As Variant
  
  FileName = Application.GetSaveAsFilename("TurtleGraphics.emf", "Enhanced Meta File (*.emf),*.emf", 1, "Save Canvas as", "Save")
  If FileName <> False Then
    If Dir(FileName) <> "" Then
      Kill FileName
    End If
    turtle.SaveCanvas (FileName), ttEMF
  End If
End Sub


Private Sub cmdRun_Click()
  YourProgramHere
End Sub

Private Sub CommandButton1_Click()

End Sub
