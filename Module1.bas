Attribute VB_Name = "Module1"
Option Explicit

Public ThetS As Single
Public ThetC As Single
Public Declare Function GdiTransparentBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Boolean
Public Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, _
                ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, _
                ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function ReleaseCapture Lib "user32.dll" () As Long

Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Public Const WM_NCLBUTTONDOWN As Integer = &HA1
Public Const WS_EX_LAYERED As Long = &H80000
Public Const GWL_EXSTYLE As Integer = (-20)
Public Const LWA_COLORKEY As Integer = &H1
Public Const LWA_ALPHA As Integer = &H2
Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1

Private Declare Function PlgBlt Lib "gdi32" (ByVal hDCDest As Long, _
                                             lpPoint As POINTS2D, _
                                             ByVal hDCSrc As Long, _
                                             ByVal nXSrc As Long, _
                                             ByVal nYSrc As Long, _
                                             ByVal nWidth As Long, _
                                             ByVal nHeight As Long, _
                                             ByVal hbmMask As Long, _
                                             ByVal xMask As Long, _
                                             ByVal yMask As Long) As Long

Public Const NotPI = 3.14159265238 / 180

Private Type POINTS2D
    X As Long
    Y As Long
End Type

Public Sub Rotate(ByRef picDestHdc As Long, xPos As Long, yPos As Long, _
                     ByVal Angle As Long, _
                     ByRef picSrcHdc As Long, srcXoffset As Long, srcYoffset As Long, _
                     ByVal srcWidth As Long, ByVal srcHeight As Long)

  Dim Points(3) As POINTS2D
  Dim DefPoints(3) As POINTS2D

  Dim Ret As Long

    Points(0).X = -srcWidth * 0.5
    Points(0).Y = -srcHeight * 0.5

    Points(1).X = Points(0).X + srcWidth
    Points(1).Y = Points(0).Y

    Points(2).X = Points(0).X
    Points(2).Y = Points(0).Y + srcHeight

    ThetS = Sin(Angle * NotPI)
    ThetC = Cos(Angle * NotPI)

    DefPoints(0).X = (Points(0).X * ThetC - Points(0).Y * ThetS) + xPos
    DefPoints(0).Y = (Points(0).X * ThetS + Points(0).Y * ThetC) + yPos

    DefPoints(1).X = (Points(1).X * ThetC - Points(1).Y * ThetS) + xPos
    DefPoints(1).Y = (Points(1).X * ThetS + Points(1).Y * ThetC) + yPos

    DefPoints(2).X = (Points(2).X * ThetC - Points(2).Y * ThetS) + xPos
    DefPoints(2).Y = (Points(2).X * ThetS + Points(2).Y * ThetC) + yPos

    PlgBlt picDestHdc, DefPoints(0), picSrcHdc, srcXoffset, srcYoffset, srcWidth, srcHeight, 0, 0, 0

End Sub

Public Sub SetOnTop(Frm As Form, OnTop As Long)

    If OnTop = -1 Then
        apiSetWindowPos Frm.hWnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE + SWP_NOSIZE
      Else
        apiSetWindowPos Frm.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOMOVE + SWP_NOSIZE
    End If

End Sub
