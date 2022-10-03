VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmEyeDropper 
   Caption         =   "Eye Dropper"
   ClientHeight    =   1290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   1800
   OleObjectBlob   =   "frmEyeDropper.frx":0000
   StartUpPosition =   1  'Centrar en propietario
End
Attribute VB_Name = "frmEyeDropper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type

Private m_blnEyeDropping As Boolean

Const HWND_TOP = 0  '//moves to top of Zorder
Const SWP_NOSIZE = &H1  '//Overwrites cx & cy to not resize window.

#If Win64 Then
 
    Private Declare PtrSafe Function FindWindowA _
        Lib "user32.dll" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr
 

 
    Private Declare PtrSafe Function GetWindowLong Lib "user32" _
                    Alias "GetWindowLongA" _
                   (ByVal hWnd As Long, _
                    ByVal nIndex As Long) As Long
    
    
        Private Declare PtrSafe Function SetWindowLong Lib "user32" _
                    Alias "SetWindowLongA" _
                   (ByVal hWnd As Long, _
                    ByVal nIndex As Long, _
                    ByVal dwNewLong As Long) As Long
    
    
        Private Declare PtrSafe Function DrawMenuBar Lib "user32" _
                   (ByVal hWnd As Long) As Long
 
 
     Private mlnghWnd As LongPtr
  
     Public Property Get hWnd() As LongPtr
         hWnd = mlnghWnd
     End Property
   
#Else
 
    Private Declare Function FindWindowA _
        Lib "user32.dll" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long
 


    Private Declare Function GetWindowLong Lib "user32" _
                Alias "GetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long) As Long


    Private Declare Function SetWindowLong Lib "user32" _
                Alias "SetWindowLongA" _
               (ByVal hWnd As Long, _
                ByVal nIndex As Long, _
                ByVal dwNewLong As Long) As Long


    Private Declare Function DrawMenuBar Lib "user32" _
               (ByVal hWnd As Long) As Long

    Private mlnghWnd As Long
 
    Public Property Get hWnd() As Long
        hWnd = mlnghWnd
    End Property

Private Sub imgColor_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub imgColor_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
  imgBorder.borderStyle = fmBorderStyleNone
End Sub

#End If

Private Sub imgDropper_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Pnt As POINTAPI
    
    If m_blnEyeDropping = True Then
        GetCursorPos Pnt
        SetWindowPos Me.hWnd, HWND_TOP, Pnt.X + 10, Pnt.Y + 10, 0, 0, SWP_NOSIZE
        imgColor.BackColor = m_ColorUnderDropper()
        imgBorder.borderStyle = fmBorderStyleNone
    Else
      imgBorder.borderStyle = fmBorderStyleSingle
    
    End If


End Sub

 
Private Sub UserForm_Initialize()
 
    StorehWnd
    HideBar Me
 
End Sub
 
Private Sub StorehWnd()
 
    Dim strCaption As String
    Dim strClass As String
 
    'class name changed in Office 2000
    If Val(Application.Version) >= 9 Then
        strClass = "ThunderDFrame"
    Else
        strClass = "ThunderXFrame"
    End If
 
    'remember the caption so we can
    'restore it when we're done
    strCaption = Me.Caption
 
    'give the userform a random
    'unique caption so we can reliably
    'get a handle to its window
    Randomize
    Me.Caption = CStr(Rnd)
 
    'store the handle so we can use
    'it for the userform's lifetime
    mlnghWnd = FindWindowA(strClass, Me.Caption)
 
    'set the caption back again
    Me.Caption = strCaption
 
End Sub

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

Private Sub imgDropper_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim Pnt As POINTAPI
    
  If Button = 1 Then
    GetCursorPos Pnt
    SetWindowPos Me.hWnd, HWND_TOP, Pnt.X + 10, Pnt.Y + 10, 0, 0, SWP_NOSIZE
    imgDropper.ZOrder 1
    m_blnEyeDropping = True
    If chkHideExcel.value = True Then
      Application.Visible = False
    End If
  End If
End Sub

Private Sub imgDropper_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 And m_blnEyeDropping Then
        'If chkHide Then Application.Visible = True
        imgDropper.ZOrder 0
        Application.Visible = True
        Me.Hide
    End If
    ' end dropping
    m_blnEyeDropping = False

End Sub


Private Function m_ColorUnderDropper() As Long
    Dim Pnt As POINTAPI
    GetCursorPos Pnt
    m_ColorUnderDropper = GetPixel(GetDC(0), Pnt.X, Pnt.Y)   ' GetDC(0) returns the screen's hdc
End Function


Sub HideBar(frm As Object)

  Dim Style As Long, Menu As Long, hWndForm As Long
  hWndForm = Me.hWnd
  Style = GetWindowLong(hWndForm, &HFFF0)
  Style = Style And Not &HC00000
  SetWindowLong hWndForm, &HFFF0, Style
  DrawMenuBar hWndForm

End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
imgBorder.borderStyle = fmBorderStyleNone
End Sub
