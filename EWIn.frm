VERSION 5.00
Begin VB.Form EDesk 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   165
   ClientTop       =   765
   ClientWidth     =   4485
   Icon            =   "EWIn.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "EWIn.frx":08CA
   ScaleHeight     =   2985
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
   Begin VB.CheckBox Sosp 
      Caption         =   "&Pause"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2520
      Width           =   1275
   End
   Begin VB.CommandButton CkLOG 
      Caption         =   "&Control Panel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   300
      TabIndex        =   5
      Top             =   2520
      Width           =   1395
   End
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   2460
      Top             =   60
   End
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1980
      Top             =   60
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Forza Lettura"
      Height          =   375
      Left            =   1380
      TabIndex        =   0
      Top             =   3000
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   3675
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   480
      Width           =   3555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1260
      Width           =   3075
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   1020
      Width           =   2955
   End
   Begin VB.Menu SysMenu 
      Caption         =   "File"
      Begin VB.Menu Open 
         Caption         =   "Open"
      End
      Begin VB.Menu Options 
         Caption         =   "Options"
      End
      Begin VB.Menu CLog 
         Caption         =   "Projects"
      End
      Begin VB.Menu Sospe 
         Caption         =   "Pause"
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu Infos 
         Caption         =   "About"
      End
      Begin VB.Menu Closei 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "EDesk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GlobalTime As Double
Dim LocalTime As Double

Sub SetParameters()

  MTimeout = Val(GetSetting(App.EXEName, "Configurazione", "TimeOut Min", 0))
  STimeout = Val(GetSetting(App.EXEName, "Configurazione", "TimeOut Sec", 10))
  NotifyDex = GetSetting(App.EXEName, "Configurazione", "Descrivi Attività", "Si") = "Si"

End Sub

Private Sub CkLOG_Click()
  VLOG.Show 1
End Sub

Private Sub CLog_Click()
CkLOG_Click
End Sub

Private Sub Closei_Click()
 On Error Resume Next
 Unload Me
 End
End Sub

Private Sub Command1_Click()
 Dim hDeskTop As Long
 hDeskTop = GetDesktopWindow()
 ProjectFound = False
 EnumChildWindows hDeskTop, AddressOf EnumeraFinestreTopLevel, 0
 
 If Not ProjectFound And Len(LPrj) > 0 Then
    UpdateProject LPrj, True   ' Chiude il Progetto Precedente
    LPrj = ""
 End If

End Sub



Private Sub Form_Load()
  ' Hide the form
    On Error Resume Next
    
    With Me
        .Top = -10000
        .Left = -10000
        .WindowState = vbMinimized
    End With
    
    On Error GoTo 0
    
    SetParameters
         
    With nidProgramData
        .cbSize = Len(nidProgramData)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        ' This is the event that will trigger when stuff happens
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "SpyIDE" & vbNullChar
    End With

    ' Call Notify...
    Shell_NotifyIcon NIM_ADD, nidProgramData


 Label1 = ""
 Label2 = ""
 Label3 = ""
 Label4 = ""
 
 GlobalTime = Timer
 Timer1_Timer
 
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error GoTo Form_MouseMove_err:
      
    ' This procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long
    
    ' The value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = x
    Else
        msg = x / Screen.TwipsPerPixelX
    End If

    
    Select Case msg
        Case WM_LBUTTONUP
    ' process single click on your icon
   
   ' Call Command1_Click
                    
        Case WM_LBUTTONDBLCLK
    ' Process double click on your icon
         Open_Click
       
                    
        Case WM_RBUTTONUP
    ' Usually display popup menu
           PopupMenu SysMenu

        Case WM_MOUSEISMOVING
    ' Do Somthing...
            
    End Select

    Exit Sub
    
Form_MouseMove_err:
    
    ' Your Error handler goes here!

End Sub


Private Sub Form_Resize()
      If WindowState = 1 Then Me.Hide
      If WindowState = 0 Then Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)

' If MsgBox("Chiudere SpyIDE?", 32 + 4, "Chiusura Spy") <> 6 Then
'    Cancel = True
'    Exit Sub
' End If
 
 If Len(LPrj) > 0 Then UpdateProject LPrj, True
    
' Rimuovi da Not.Try

 Shell_NotifyIcon NIM_DELETE, nidProgramData
  
End Sub


Private Sub Infos_Click()
'  Info.Show 1
  Link1 = "INFO"
 Splash.Show
End Sub

Private Sub Open_Click()
          Me.WindowState = vbNormal
          Me.Show
          iTop = (Screen.Height - Me.Height) \ 2
          iLeft = (Screen.Width - Me.Width) \ 2
    
        'If iTop And iLeft Then
          Me.Move iLeft, iTop

End Sub

Private Sub Options_Click()
 OPts.Show 1
End Sub

Private Sub Sosp_Click()

Sospe.Checked = Sosp = 1

If Sosp = 1 Then
 Timer1.Enabled = False
 Timer2.Enabled = False
 Caption = "Paused from the user"
 nidProgramData.szTip = Caption & vbNullChar
 Shell_NotifyIcon NIM_MODIFY, nidProgramData
Else
 Timer1.Enabled = True
 Timer2.Enabled = True
End If
 
End Sub

Private Sub Sospe_Click()
 Sosp = Abs(Sosp = 0)
End Sub

Private Sub Timer1_Timer()
 
 Dim hDeskTop As Long
 hDeskTop = GetDesktopWindow()
 ProjectFound = False
 EnumChildWindows hDeskTop, AddressOf EnumeraFinestreTopLevel, 0
 
 If ProjectFound Then
    Timer2.Enabled = True
 Else
    
    If Not ProjectFound And Len(LPrj) > 0 Then
       UpdateProject LPrj, True   ' closes the previous project
       LPrj = ""
    End If
    
    Label3 = "Inactive."
    Label4 = ""
    Label1 = ""
    Label2 = ""
    Caption = Label3
    nidProgramData.szTip = Caption & vbNullChar
    Shell_NotifyIcon NIM_MODIFY, nidProgramData
    Timer2.Enabled = False
    
 End If
 
End Sub




Private Sub Timer2_Timer()
  Static Oj As String
  
  If LPrj <> Oj Then
     Oj = LPrj
     LocalTime = Timer
  End If
  
  Gt = CLng(Timer - GlobalTime)
  Lt = CLng(Timer - LocalTime)
  
  k = Gt
  GoSub BldTime
  Label3 = "Spy Activity: " & St
  k = Lt
  GoSub BldTime
  Label4 = "Current project Activity: " & St
  
  Caption = LPrj & " (" & NLog & ") - " & St
  nidProgramData.szTip = Caption & vbNullChar
  Shell_NotifyIcon NIM_MODIFY, nidProgramData
  Exit Sub
  
BldTime:
     
     m = Fix(k / 60)
     h = Fix(m / 60)
     If h > 0 Then m = m Mod 60
     s = k Mod 60
     St = Format$(h, "0#") & ":" & Format$(m, "0#") & ":" & Format$(s, "0#")
     
     OrePub = h
     MinPub = m
     SecPub = s
     
Return
  
End Sub


