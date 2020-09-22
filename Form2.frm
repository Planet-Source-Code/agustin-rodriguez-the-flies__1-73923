VERSION 5.00
Begin VB.Form The_Fly 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   ClientHeight    =   1275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1275
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   85
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   85
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   600
   End
End
Attribute VB_Name = "The_Fly"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Angle As Long
Private Adjuste As Integer
Private Velocity As Integer
Private Stopped As Boolean
Private Direction As Integer
Private PX As Single
Private PY As Single

Private Sub Form_DblClick()

  Dim X As New The_Fly

    X.Show
    X.Move Left + 10, Top + 10

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 27 Then
        Main_Frm.End_Program
    End If

End Sub

Private Sub Form_Load()

  Dim Ret As Long

    Adjuste = 25
    Velocity = 60
    Angle = Int(Rnd * 360)
    Stopped = 0
    Direction = 1
    Set_Trajectory

    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes hWnd, 255, 255, LWA_COLORKEY Or LWA_ALPHA
    Move Int(Rnd * Screen.Width), Int(Rnd * Screen.Height)
    SetOnTop Me, True

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2&, 0&

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Stopped = True

End Sub

Private Sub Timer1_Timer()

  Static X As Integer

    If Stopped Then
        Main_Frm.Picture2.Picture = Main_Frm.FlyStopted(X).Picture
      Else
        Main_Frm.Picture2.Picture = Main_Frm.Fly(X).Picture
    End If

    Rotate Main_Frm.Picture1.hDC, 42, 42, Angle + Adjuste, Main_Frm.Picture2.hDC, 0, 0, 60, 60
    Main_Frm.Picture1.Refresh
    Picture = Main_Frm.Picture1.Image
    Refresh
    X = X + Direction

    If X = 14 Or X = 0 Then
        Direction = Direction * -1
    End If

    If Stopped = 0 Then
        Move Left + PX, Top + PY
    End If

    If Left + Width < 0 Then
        Left = Screen.Width - 5
    End If
    If Left > Screen.Width Then
        Left = 5
    End If
    If Top + Height < 0 Then
        Top = Screen.Height - 5
    End If
    If Top > Screen.Height Then
        Top = 5
    End If

    If Int(Rnd * 3) = 1 And Stopped = False Then
        Angle = Angle + IIf(Int(Rnd * 2) = 1, 1, -1) * Int(Rnd * 10)
        Set_Trajectory
        SetOnTop Me, True
        Velocity = Int(Rnd * 100) + 10
    End If

    If Int(Rnd * 10) = 1 Then
        Stopped = Stopped Xor -1
        Adjuste = IIf(Stopped, -114, 25)
    End If

End Sub

Private Sub Set_Trajectory()

    PY = Sin((Angle) * NotPI) * Velocity
    PX = Cos((Angle) * NotPI) * Velocity

End Sub
