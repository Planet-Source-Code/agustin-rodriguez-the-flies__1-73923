VERSION 5.00
Begin VB.Form Main_Frm 
   AutoRedraw      =   -1  'True
   BackColor       =   &H000000FF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2385
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6090
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   159
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   406
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Background 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   3525
      Left            =   6840
      Picture         =   "Form1.frx":164A
      ScaleHeight     =   235
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   34
      Top             =   120
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.PictureBox Auxiliar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   353
      TabIndex        =   33
      Top             =   2520
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.PictureBox Splash_Pic 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   2385
      Left            =   0
      ScaleHeight     =   159
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   400
      TabIndex        =   32
      Top             =   2280
      Visible         =   0   'False
      Width           =   6000
   End
   Begin VB.Timer Splash 
      Interval        =   100
      Left            =   840
      Top             =   6240
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   14
      Left            =   9360
      Picture         =   "Form1.frx":4641C
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   31
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   13
      Left            =   8400
      Picture         =   "Form1.frx":48E8E
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   30
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   12
      Left            =   7440
      Picture         =   "Form1.frx":4B900
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   29
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   11
      Left            =   6480
      Picture         =   "Form1.frx":4E372
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   28
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   10
      Left            =   5520
      Picture         =   "Form1.frx":50DE4
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   27
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   9
      Left            =   9360
      Picture         =   "Form1.frx":53856
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   26
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   8
      Left            =   8400
      Picture         =   "Form1.frx":562C8
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   25
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   7
      Left            =   7440
      Picture         =   "Form1.frx":58D3A
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   24
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   6
      Left            =   6480
      Picture         =   "Form1.frx":5B7AC
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   23
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   5
      Left            =   5520
      Picture         =   "Form1.frx":5E21E
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   22
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   4
      Left            =   9360
      Picture         =   "Form1.frx":60C90
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   21
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   3
      Left            =   8400
      Picture         =   "Form1.frx":63702
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   20
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   2
      Left            =   7440
      Picture         =   "Form1.frx":66174
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   19
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   1
      Left            =   6480
      Picture         =   "Form1.frx":68BE6
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   18
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox FlyStopted 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   0
      Left            =   5520
      Picture         =   "Form1.frx":6B658
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   17
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   14
      Left            =   4080
      Picture         =   "Form1.frx":6E0CA
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   16
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   13
      Left            =   3120
      Picture         =   "Form1.frx":70B3C
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   15
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   12
      Left            =   2160
      Picture         =   "Form1.frx":735AE
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   14
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   11
      Left            =   1200
      Picture         =   "Form1.frx":76020
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   13
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   10
      Left            =   240
      Picture         =   "Form1.frx":78A92
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   12
      Top             =   4800
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   9
      Left            =   4080
      Picture         =   "Form1.frx":7B504
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   11
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1272
      Left            =   7200
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   10
      Top             =   5760
      Visible         =   0   'False
      Width           =   1272
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1272
      Left            =   1920
      ScaleHeight     =   85
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   9
      Top             =   5760
      Width           =   1272
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   0
      Left            =   240
      Picture         =   "Form1.frx":7DF76
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   8
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   1
      Left            =   1200
      Picture         =   "Form1.frx":809E8
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   7
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   2
      Left            =   2160
      Picture         =   "Form1.frx":8345A
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   6
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   3
      Left            =   3120
      Picture         =   "Form1.frx":85ECC
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   5
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   4
      Left            =   4080
      Picture         =   "Form1.frx":8893E
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   4
      Top             =   2880
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   5
      Left            =   240
      Picture         =   "Form1.frx":8B3B0
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   6
      Left            =   1200
      Picture         =   "Form1.frx":8DE22
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   7
      Left            =   2160
      Picture         =   "Form1.frx":90894
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   1
      Top             =   3840
      Width           =   900
   End
   Begin VB.PictureBox Fly 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   900
      Index           =   8
      Left            =   3120
      Picture         =   "Form1.frx":93306
      ScaleHeight     =   60
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   60
      TabIndex        =   0
      Top             =   3840
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click The Splash Screen to Hide it"
      Height          =   195
      Index           =   2
      Left            =   2520
      TabIndex        =   38
      Top             =   600
      Width           =   2985
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Double Click on any Fly  to Add"
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   37
      Top             =   360
      Width           =   2220
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "ESC to EXIT"
      Height          =   195
      Index           =   0
      Left            =   3480
      TabIndex        =   35
      Top             =   120
      Width           =   900
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      Height          =   855
      Left            =   2400
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   3255
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "  X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5640
      TabIndex        =   36
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H0000C0C0&
      BackStyle       =   1  'Opaque
      Height          =   375
      Left            =   5520
      Shape           =   4  'Rounded Rectangle
      Top             =   0
      Width           =   495
   End
End
Attribute VB_Name = "Main_Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_DblClick()
Splash.Enabled = False
Hide
End Sub

Private Sub Form_Load()

  Dim I As Integer
  Dim Ret As Long
  
  Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
  Ret = Ret Or WS_EX_LAYERED
  SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
  SetLayeredWindowAttributes hWnd, 255, 155, LWA_COLORKEY Or LWA_ALPHA
  
  Show_Text
  Dim Fly(0 To 9) As New The_Fly

    For I = 0 To 9
        Fly(I).Show
    Next I

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, WM_NCLBUTTONDOWN, 2&, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)

  Dim Frm As Form

    For Each Frm In Forms
        Unload Frm
    Next Frm

    Unload The_Fly

End Sub

Public Sub End_Program()

    Unload Me

End Sub

Private Sub Label2_Click()
    
    Label1(0).Visible = False
    Label1(1).Visible = False
    Label1(2).Visible = False
    Label2.Visible = False
    Shape1.Visible = False
    Shape2.Visible = False
    
End Sub

Private Sub Splash_Timer()

Const Run_step As Integer = 3

GdiTransparentBlt Background.hDC, 0, 0, Splash_Pic.ScaleWidth, Splash_Pic.ScaleHeight, Splash_Pic.hDC, 0, 0, Splash_Pic.ScaleWidth, Splash_Pic.ScaleHeight, vbYellow
Background.Refresh
BitBlt hDC, 0, 0, Splash_Pic.Width, Splash_Pic.Height, Background.hDC, 0, 0, vbSrcCopy
Background.Cls
BitBlt Auxiliar.hDC, 0, 0, Background.ScaleWidth, Run_step, Background.hDC, 0, 0, vbSrcCopy
BitBlt Background.hDC, 0, 0, Background.ScaleWidth, Background.ScaleHeight - Run_step, Background.hDC, 0, Run_step, vbSrcCopy
BitBlt Background.hDC, 0, Background.ScaleHeight - Run_step, Background.ScaleWidth, Run_step, Auxiliar.hDC, 0, 0, vbSrcCopy
Refresh
Background.Picture = Background.Image

End Sub
Private Sub Show_Text()
With Splash_Pic
    .FontName = "Arial Black"
    .FontSize = 50
    
    .CurrentX = 0
    .CurrentY = -20
    .ForeColor = vbBlack
    Splash_Pic.Print "THE"
    
    .CurrentX = 3
    .CurrentY = -19
    .ForeColor = vbYellow
    Splash_Pic.Print "THE"
    
    .FontSize = 80
    .CurrentX = 50
    .CurrentY = 20
    .ForeColor = vbBlack
    Splash_Pic.Print "FLIES"
    
    .CurrentX = 55
    .CurrentY = 25
    .ForeColor = vbYellow
    Splash_Pic.Print "FLIES"
End With

Copyright
End Sub
Private Sub Copyright()
With Splash_Pic
    .FontSize = 10
    .CurrentX = 210
    .CurrentY = 140
    .ForeColor = vbWhite
    Splash_Pic.Print "© Agustin Rodriguez"
    .CurrentX = 211
    .CurrentY = 141
    .ForeColor = vbBlack
    Splash_Pic.Print "© Agustin Rodriguez"
    End With
End Sub
