VERSION 5.00
Begin VB.Form OPts 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3840
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   420
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3240
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   420
      Width           =   375
   End
   Begin VB.CheckBox RDex 
      Alignment       =   1  'Right Justify
      Caption         =   "Describes the activity after project changes"
      Height          =   315
      Left            =   720
      TabIndex        =   2
      Top             =   60
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1020
      Width           =   795
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   2340
      TabIndex        =   0
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "s"
      Height          =   195
      Left            =   4260
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.Label Label2 
      Caption         =   "m"
      Height          =   195
      Left            =   3660
      TabIndex        =   5
      Top             =   480
      Width           =   195
   End
   Begin VB.Label Label1 
      Caption         =   "Do not LOG information for time less than"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   3075
   End
End
Attribute VB_Name = "OPts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
  
  If Val(Text1) > 59 Or Val(Text2) > 59 Then
     MsgBox "Wrong parameters", 16
     Exit Sub
  End If
  
  MTimeout = Val(Text1)
  STimeout = Val(Text2)
  NotifyDex = RDex = 1
  
  If NotifyDex Then
     ND$ = "Si"
  Else
     ND$ = "No"
  End If
  
  SaveSetting App.EXEName, "Configurazione", "TimeOut Min", Val(Text1)
  SaveSetting App.EXEName, "Configurazione", "TimeOut Sec", Val(Text2)
  SaveSetting App.EXEName, "Configurazione", "Descrivi Attività", ND$

  Unload Me
 
End Sub

Private Sub Form_Load()
  
  Text1 = MTimeout
  Text2 = STimeout
  RDex = Abs(NotifyDex)
  
End Sub


