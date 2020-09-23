VERSION 5.00
Begin VB.Form VLOG 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "History Projects"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   120
      TabIndex        =   2
      Top             =   3420
      Width           =   5775
      Begin VB.Label Label1 
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   60
         TabIndex        =   3
         Top             =   240
         Width           =   5595
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      Left            =   120
      TabIndex        =   1
      Top             =   540
      Width           =   5775
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   5775
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label2"
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   4740
      Width           =   5775
   End
End
Attribute VB_Name = "VLOG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim FIni As String

Private Sub Combo1_Click()
  Dim Sez As String
  Dim Key As String
  Dim St As String
  Dim RkL As RkT
  Dim th As Integer
  Dim tm As Integer
  
  List1.Clear
  
  
 If Combo1.ListIndex > 0 Then
    Prj = Combo1
    Sez = Prj
    NL = Val(ReadIni(Sez, "Next LOG", FIni))
    For i = 0 To NL - 1
       Key = Str(i)
       St = ReadIni(Sez, Key, FIni)
       RkL = Scompone(St)
       If RkL.Tempo <> "00:00" Then
          List1.AddItem RkL.DataI & Chr(9) & RkL.OraI & Chr(9) & RkL.Tempo & Chr(9) & RkL.Dex
          GoSub Contab
       End If
    Next
 ElseIf Combo1.ListIndex = 0 Then
 
    For j = 1 To Combo1.ListCount - 1
          Sez = Combo1.List(j)
          NL = Val(ReadIni(Sez, "Next LOG", FIni))
          For i = 0 To NL - 1
               Key = Str(i)
               St = ReadIni(Sez, Key, FIni)
               RkL = Scompone(St)
               If RkL.Tempo <> "00:00" Then
                    List1.AddItem RkL.DataI & Chr(9) & RkL.OraI & Chr(9) & RkL.Tempo & Chr(9) & RkL.Dex
                    GoSub Contab
               End If
          Next i
    Next j
    
 End If
 
 
 Label2 = " Total time : " & Format$(th, "00") & " hours, " & Format$(tm, "00") & " min. "
 
 
 Exit Sub

Contab:

' Sum

h = Val(Mid$(RkL.Tempo, 1, 2))
m = Val(Mid$(RkL.Tempo, 4, 2))

th = th + h
tm = tm + m
If tm > 59 Then
   tm = tm - 59
   th = th + 1
End If
Return

End Sub


Private Sub Form_Load()
  Dim St As String
  Label1 = ""
  Label2 = ""
  FIni = App.Path & "\VB5PRJ.LOG"
  Combo1.AddItem "All projects"
  
  Open FIni For Input As #1
  
  Do Until EOF(1)
    Line Input #1, St
    If Mid$(St, 1, 1) = "[" Then
       St = Mid$(St, 2, Len(St) - 2)
       Combo1.AddItem St
    End If
  Loop
    
 Close #1

 FX_ExplodeForm Me, 500


End Sub


Private Sub Form_Unload(Cancel As Integer)
 FX_ImplodeForm Me, 2, 300, 1
End Sub


Private Sub List1_Click()
 If List1.ListIndex < 0 Then Exit Sub
 St = List1.List(List1.ListIndex)
 
 Ps = 0
 Do
   Ps = InStr(Ps + 1, St, Chr(9))
   If Ps > 0 Then
      Bs = Ps
   Else
      Exit Do
   End If
 Loop
 
 If Bs > 0 Then Label1 = Mid$(St, Bs + 1, Len(St))
 
 
End Sub


