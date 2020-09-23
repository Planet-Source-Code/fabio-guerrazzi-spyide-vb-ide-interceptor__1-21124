Attribute VB_Name = "Desktop"
Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function EnumDesktopWindows Lib "user32" (ByVal hDeskTop As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
Declare Function GetDesktopWindow Lib "user32" () As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long


Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetClassWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long


Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public LPrj As String ' Ultimo Progetto in memoria
Public UCont As Long ' Numero Letture eseguite su progetti validi
Public ProjectFound As Boolean
Public DexOpen As Boolean
Public DexTxt As String

Public Link1 As String
Public NLog As Long

Type RkT
  DataI As String
  OraI As String
  DataF As String
  OraF As String
  Tempo As String
  Dex As String
End Type

Public MTimeout As Byte
Public STimeout As Byte
Public NotifyDex As Boolean

Public OrePub As Integer
Public MinPub As Integer
Public SecPub As Integer


Public Function EnumeraFinestreTopLevel(ByVal hwndTopLevel As Long, lParam As Long) As Long
    Dim Testo As String, NumCar As Integer
    Dim Class As String, NumClassCar As Integer
    Dim IsOk As Boolean

' Loop per tutte le finestre attive:
' GetWidowTExt Pone in "Testo" la caption relativa della finestra
' GetClassName Pone in "Class" in nome della Classe
' Se Class non è nullo controlla se è la classe ThunderMain (Finestra Main di progettazione VB)
' Se si Testo contiene il nome del progetto VB


    Testo = Space$(250)
    Class = Space$(250)
    NumCar = GetWindowText(hwndTopLevel, Testo, 250)
    NumClassCar = GetClassName(hwndTopLevel, Class, 250)
    
    If NumClassCar > 0 Then
       Class = Trim$(Class)
       Class = Mid$(Class, 1, Len(Class) - 1)
       If NumCar > 0 Then
            If Class = "ThunderMain" Or Class = "PROJECT" Then
              
               Testo = Trim$(Testo)
               Testo = Mid$(Testo, 1, Len(Testo) - 1)
               If InStr(Testo, "Microsoft Visual Basic") = 0 And Testo <> "Progetto1" And Testo <> "Project1" And InStr(Testo, "Progetto -") = 0 Then
                  If Class = "PROJECT" Then Testo = Testo & " (16)"
                  If Len(LPrj) > 0 And LPrj <> Testo Then
                     UpdateProject LPrj, True   ' Chiude il Progetto Precedente
                  End If
                  ProjectFound = True
                  UpdateProject Testo, False ' Apre
                  LPrj = Testo
               End If
            End If
       End If
    End If
    
    
    EnumeraFinestreTopLevel = True
    
End Function

Function Scompone(St As String) As RkT
  If Len(St) = 0 Then Exit Function
  Ps = 0
  b = 0
  Do
    Ps = InStr(Ps + 1, St, "|")
    If Ps = 0 Then
       Exit Do
    Else
       b = b + 1
       s = Mid$(St, Bs + 1, Ps - (Bs + 1))
       Select Case b
         Case 1
           Scompone.DataI = s
         Case 2
           Scompone.OraI = s
         Case 3
           Scompone.DataF = s
         Case 4
           Scompone.OraF = s
         Case 5
           Scompone.Tempo = s
  '       Case 6
  '         Scompone.Dex = s
       End Select
       
       Bs = Ps
    End If
    
  Loop


  Scompone.Dex = Mid$(St, Bs + 1, Len(St))

  


End Function

Sub UpdateProject(Prj As String, FClose As Boolean)

    Dim FIni As String
    Dim Sez As String
    Dim DataL As String
    Dim Key As String
    Dim St As String
    Dim Stime As Double
    Dim RkL As RkT
    
    UCont = UCont + 1
    FIni = App.Path & "\VB5PRJ.LOG"
    NLog = Val(ReadIni(Prj, "Next LOG", FIni))
    Key = CStr(NLog)
    Sez = Prj
    
    St = ReadIni(Sez, Key, FIni)
    RkL = Scompone(St)
  
' closing SPYIDE or project was changed
    
    If FClose Then
        
        GoSub Tempo
        
        hh = OrePub
        mm = MinPub
        
        mP = (MinPub * 60) + SecPub
        mO = (MTimeout * 60) + STimeout
        
        If mP <= mO Then ' not enought time spent to log
           Call WriteINI(Sez, Key, " ", FIni)
           Exit Sub
        End If
        
        If NotifyDex Then
             Link1 = ""
             FDex.Show
             FDex.Label3 = Prj
             FDex.Label2 = "Elapsed time: " & EDesk.Label2 '& hh & " hours e " & mm & " minuts."
        
             Do While DexOpen = True
                DoEvents
                Beep
                Stime = Timer
                GoSub Wait20Second
             Loop
        
             If Link1 = "Cancel" Then
                 Call WriteINI(Sez, Key, " ", FIni)
                 Exit Sub
             End If
        
             St = DexTxt
         Else
             St = "** No Dex"
         End If
        
        RkL.Tempo = TimesL
        RkL.DataF = Format$(Now, "dd/mm/yyyy")
        RkL.OraF = Format$(Now, "hh:mm")
        RkL.Dex = St
        
        Call WriteINI(Sez, "Next LOG", NLog + 1, FIni)
        Call WriteINI(Sez, "Tempo Totale", TimesT, FIni)
        GoSub Scrivi
        
        Exit Sub
    
    End If
    
' Lettura Iniziale del log
    
    DataL = RkL.DataI
    OraL = RkL.OraI
    
    
' Se è la prima lettura o la precedente non si è conclusa bene scrive i valori di partenza o azzera i precedenti
    
    If Len(DataL) = 0 Or LPrj <> Prj And RkL.Dex = "****** READING *******" Then
       
       If Not FileExists(FIni) Then
          ' Crea le note
          Open FIni For Output As #1
            Print #1, ";SpyIDE v1.0 - 1998/2001 (c) Fabio Guerrazzi"
            Print #1, " "
            Print #1, ";Tabs meaning: (|)"
            Print #1, ";1-Date of opening project"
            Print #1, ";2-Hour of opening project"
            Print #1, ";3-Date of closing project"
            Print #1, ";4-Hour of closing project"
            Print #1, ";5-Elapsed time"
            Print #1, ";6-Description"
          Close
       End If
       
       RkL.DataI = Format$(Now, "dd/mm/yyyy")
       RkL.OraI = Format$(Now, "hh:mm")
       RkL.Tempo = "00:00"
       RkL.Dex = "****** READING *******"
       RkL.DataF = RkL.DataI
       RkL.OraF = RkL.OraI
       EDesk.Label1 = CStr(NLog) & "Th LOG started at " & RkL.OraI
       GoSub Scrivi
       Exit Sub
       
    End If
    
    GoSub Tempo

Exit Sub

Wait20Second:
  
  Do Until Timer - Stime > 20
    If Not DexOpen Then Exit Do
    DoEvents
  Loop

Return


Tempo:
    
   
       TempoT = ReadIni(Prj, "Tempo Totale", FIni)
       
    '   OreL = Val(Mid$(RkL.OraI, 1, 2))
    '   MinL = Val(Mid$(RkL.OraI, 4, 2))

    '   Adesso$ = Format(Now, "hh:mm")
    
    '   OreA = Val(Mid$(Adesso$, 1, 2))
    '   MinA = Val(Mid$(Adesso$, 4, 2))
    
       OreT = Val(Mid$(TempoT, 1, 2))
       MinT = Val(Mid$(TempoT, 4, 2))
    
    '   HDif = OreA - OreL
    '   MDif = MinA - MinL
    
    '   If MDif <= 0 And HDif > 0 Then
    '      HDif = HDif - 1
    '      MDif = Abs(MDif)
    '   End If
       
    HDif = OrePub
    MDif = MinPub
    
       OreT = OreT + HDif
       MinT = MinT + MDif
    
       If MinT > 59 Then
          OreT = OreT + 1
          MinT = MinT - 59
       End If
    
       TimesT = Format$(OreT, "0#") & ":" & Format$(MinT, "0#")
       TimesL = Format$(HDif, "0#") & ":" & Format$(MDif, "0#")
       
    EDesk.Label2 = " " & Format$(HDif, "0#") & " Hours, " & Format$(MDif, "0#") & " mins, " & Format$(SecPub, "0#") & " sec."

Return

Scrivi:
   
   St = RkL.DataI & "|" & RkL.OraI & "|" & RkL.DataF & "|" & RkL.OraF & "|" & RkL.Tempo & "|" & RkL.Dex
  ' MsgBox "Scrittura di: " & St
   Call WriteINI(Sez, Key, St, FIni)
Return

End Sub

Function FileExists(FileName As String) As Integer
Dim i As Integer
On Error Resume Next
i = Len(Dir$(FileName))
If Err Or i = 0 Then
    FileExists = False
  Else
    FileExists = True
End If

End Function

Function ReadIni(Appname As String, KeyName As String, FileName As String) As String
  Dim sRet As String
  sRet = String(255, Chr(0))
  ReadIni = Left(sRet, GetPrivateProfileString(Appname, ByVal KeyName, "", sRet, Len(sRet), FileName))
End Function

Function WriteINI(Appname As String, KeyName As String, NewString As Variant, FileName As String) As Integer
     Dim s As Integer
     s = WritePrivateProfileString(Appname, KeyName, CStr(NewString), FileName)
     If s = False Then
         MsgBox "error writing the LOG file.", 16
         End
     End If
End Function

