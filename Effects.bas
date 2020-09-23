Attribute VB_Name = "Effects"
Option Explicit

Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal x3 As Long, ByVal y3 As Long) As Long

' Tipi e Costanti per FX_ImplodeForm e FX_ExplodeForm

Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'===============================================================================================
' Usata da FX_ImplodeForm e FX_ExplodeForm

Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "user32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long



Public Sub FX_ImplodeForm(f As Form, Direction As Integer, Movement As Integer, ModalState As Integer)
'****************************************************************
'*Author: Carl Slutter
'*
'*Description:
'*The larger the "Movement" value, the slower the "Implosion"
'*
'*Creation Date: Thursday  23 January 1997  2:42 pm
'*Revision Date: Thursday  23 January 1997  2:42 pm
'*
'*Version Number: 1.00
'
'
' Si richiama con:
'
'    Call ImplodeForm(Me, 2, 500, 1)
'
'
'
'****************************************************************
    
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, x%, y%, Cx%, Cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
        For i = Movement To 1 Step -1
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        x = myRect.Left + (formWidth - Cx) / 2
        y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, x, y, x + Cx, y + Cy
    Next i
    
    x = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
        
End Sub

Sub FX_ExplodeForm(f As Form, Movement As Integer)
    
    
    Dim myRect As RECT
    Dim formWidth%, formHeight%, i%, x%, y%, Cx%, Cy%
    Dim TheScreen As Long
    Dim Brush As Long
    
    GetWindowRect f.hwnd, myRect
    formWidth = (myRect.Right - myRect.Left)
    formHeight = myRect.Bottom - myRect.Top
    TheScreen = GetDC(0)
    Brush = CreateSolidBrush(f.BackColor)
    
    For i = 1 To Movement
        Cx = formWidth * (i / Movement)
        Cy = formHeight * (i / Movement)
        x = myRect.Left + (formWidth - Cx) / 2
        y = myRect.Top + (formHeight - Cy) / 2
        Rectangle TheScreen, x, y, x + Cx, y + Cy
    Next i
    
    x = ReleaseDC(0, TheScreen)
    DeleteObject (Brush)
    
End Sub



Sub Rounded(f As Form)
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim x3 As Long, y3 As Long
    
    Dim Ret1 As Long, Ret2 As Long
    
    X1 = 0: Y1 = 0
    X2 = f.Width / Screen.TwipsPerPixelX
    Y2 = f.Height / Screen.TwipsPerPixelY
  
    x3 = 60: y3 = 60
    
    Ret1 = CreateRoundRectRgn(X1, Y1, X2, Y2, x3, y3)
    Ret2 = SetWindowRgn(f.hwnd, Ret1, True)

End Sub

Sub Elliptic(f As Form)
    Dim X1 As Long, Y1 As Long
    Dim X2 As Long, Y2 As Long
    Dim Ret1 As Long, Ret2 As Long
    
    X1 = 0
    Y1 = 0
    X2 = f.Width / Screen.TwipsPerPixelX
    Y2 = f.Height / Screen.TwipsPerPixelY
    
    Ret1 = CreateEllipticRgn(X1, Y1, X2, Y2)
    Ret2 = SetWindowRgn(f.hwnd, Ret1, True)

End Sub

