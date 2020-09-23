VERSION 5.00
Begin VB.Form Info 
   BackColor       =   &H8000000A&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Informazioni su SpyIDE"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   2760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   120
      Top             =   720
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000008&
      Height          =   1155
      Left            =   720
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   129
      TabIndex        =   0
      Top             =   120
      Width           =   1995
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
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "Info.frx":0000
      Top             =   420
      Width           =   480
   End
End
Attribute VB_Name = "Info"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
 Label1 = "SpyIDE v1.0 (06/98) (c)Fabio Guerrazzi via Garibaldi, 52 Colle di val d'Elsa - Siena - Tel. 0577/924417 fabiog@si.tdnet.it  Il programma è Freeware e DEVE essere duplicato così com'è, senza alterarne la forma o il contenuto. Non è possibile utilizzare questo programma a scopo di lucro."
End Sub


Private Sub Timer1_Timer()
 Label1.Top = Label1.Top - 1
 If Abs(Label1.Top) >= Label1.Height Then Label1.Top = Picture1.ScaleHeight + 5
End Sub


