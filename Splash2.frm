VERSION 5.00
Begin VB.Form Splash 
   BorderStyle     =   0  'None
   ClientHeight    =   3000
   ClientLeft      =   3420
   ClientTop       =   3015
   ClientWidth     =   3000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Splash2.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   3000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1620
      Top             =   2460
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   915
      Left            =   300
      ScaleHeight     =   57
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   2415
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   2955
         Left            =   60
         TabIndex        =   1
         Top             =   1380
         Width           =   2295
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2460
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   120
      Top             =   2400
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_Description = "Startup Screen"
Option Explicit







Private Sub Form_Click()
 If Link1 = "INFO" Then
    Timer3.Enabled = False
    Unload Me
 End If
End Sub

Private Sub Form_Load()
 If Link1 = "INFO" Then
   Label1 = "SpyIDE v1.0 (98/01) (c) Fabio Guerrazzi http://digilander.io.it/WarZi/default.htm e-mails: fabiog2@libero.it fabiog@si.tdnet.it ICQ UIN: 64649187 -"
   Picture1.Visible = True
   Timer3.Enabled = True
   Timer1.Enabled = False
 Else
   Timer1.Enabled = True
 End If

End Sub

Private Sub Form_Resize()
 ' Rounded Me
 Elliptic Me
End Sub


Private Sub Timer1_Timer()
 EDesk.Show
 Timer1.Enabled = False
' FX_ImplodeForm Me, 2, 500, 0
 Unload Me
 Set Splash = Nothing
End Sub


Private Sub Timer2_Timer()
 Static bg As Long, x As Long
' x = ScaleWidth / 2
 bg = bg + 4
 DrawWidth = bg
 PSet (2300, 600), QBColor(15)
 If bg >= 10 Then
    Timer2.Enabled = False
    Cls
 End If
 DoEvents
End Sub


Private Sub Timer3_Timer()
 Label1.Top = Label1.Top - 1
 If Abs(Label1.Top) >= Label1.Height Then Label1.Top = Picture1.ScaleHeight + 5

End Sub


