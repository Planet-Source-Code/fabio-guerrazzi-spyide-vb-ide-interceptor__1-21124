Attribute VB_Name = "COMMON32"

' Utility varie
' Dichiarazioni API, Type, Const

Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Declare Function WNetAddConnection Lib "mpr.dll" Alias "WNetAddConnectionA" (ByVal lpszNetPath As String, ByVal lpszPassword As String, ByVal lpszLocalName As String) As Long
Declare Function WNetCancelConnection Lib "mpr.dll" Alias "WNetCancelConnectionA" (ByVal lpszName As String, ByVal bForce As Long) As Long
Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" (ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

'  dwPlatformId defines:
'
Public Const VER_PLATFORM_WIN32s = 0
Public Const VER_PLATFORM_WIN32_WINDOWS = 1
Public Const VER_PLATFORM_WIN32_NT = 2

Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByVal lpVersionInformation As OSVERSIONINFO) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function SetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal wNewWord As Long) As Long

Declare Function ReleaseCapture Lib "user32" () As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Integer, ByVal lParam As Long) As Long

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long


Sub Del(f As String)
  On Error Resume Next
  Kill f
  On Error GoTo 0
End Sub


Function ExtractFileFromString(Fn As String) As String
    Dim Ps As Byte
    Dim Bps As Byte
    
    Ps = 0
    Do
     Bps = Ps
     Ps = InStr(Ps + 1, Fn, "\")
    Loop While Ps > 0
    
    If Bps > 0 Then
       ExtractFileFromString = Mid$(Fn, Bps + 1, Len(Fn))
    Else
       ExtractFileFromString = Fn
    End If

End Function

Function ExtractPathFromString(St As String) As String

' Ritorna la sola path contenuta in una stringa

    Dim Ps As Byte
    Dim Bps As Byte
    
    Ps = 0
    Do
     Bps = Ps
     Ps = InStr(Ps + 1, St, "\")
    Loop While Ps > 0
    
    If Bps > 0 Then
       ExtractPathFromString = Mid$(St, 1, Bps - 1) ', Len(St))
    Else
       ExtractPathFromString = ""
    End If


End Function


Function Pad(St As Variant, l As Integer) As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT87
 
 
 If Len(St) > l Then
    Pad = Mid$(St, 1, l)
 ElseIf l > Len(St) Then
    Pad = St & Space(l - Len(St))
 Else
    Pad = St
 End If


Exit Function
 
ErrT87:
 
    RtCodeError = GestErr(Err, "Pad")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Function Crypt(St As String) As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT89
 
  
   Dim i As Integer, t As String
  ' Cripta / Decripta una stringa su base Xor 20
  ' Eseguire la prima volta per criptare la stringa
  ' ed eseguire una seconda volta per decriptarla
    
  ' NB.: La dimensione della stringa rimane invariata
    t = St
    For i = 1 To Len(t)
        Mid$(t, i, 1) = Chr$(Asc(Mid$(t, i, 1)) Xor 20)
    Next i
    Crypt = t

Exit Function
 
ErrT89:
 
    RtCodeError = GestErr(Err, "Crypt")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function



Function EstraiNomeFile(St As String) As String
Dim RtCodeError As Integer
 ' Elimina l'estensione ad un nome di file
 Dim St1 As String
On Error GoTo ErrT90
 
  Dim Ps As Integer
  
  St1 = ExtractFileFromString(St) ' Toglie la Path
  
  Ps = InStr(St1, ".")
  If Ps > 0 Then
     EstraiNomeFile = Mid$(St1, 1, Ps - 1)
  Else
     EstraiNomeFile = St1
  End If

Exit Function
 
ErrT90:
 
    RtCodeError = GestErr(Err, "EstraiNomeFile")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Function Test_Codice(UserName As String, UserNumber As Long, View As Integer) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT92
 

For i% = 1 To Len(UserName)
   Zeichen = Mid$(UserName, i%, 1)
   pchar% = Asc(Zeichen)
   TestCode = (TestCode + pchar%) + 2 * (pchar% + 7) + 343
Next i%


If View = 1 Then
   MsgBox Str(TestCode)
End If

'Now compare the built number to number read from the INI file:


If Format$(TestCode) = Format$(UserNumber) Then
   Test_Codice = True   ' Licenza trovata
Else
   Test_Codice = False    ' Licenza non trovata
End If


Exit Function
 
ErrT92:
 
    RtCodeError = GestErr(Err, "Test_Codice")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function



Function Up(KeyAscii As Integer) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT93
 

  If KeyAscii >= 94 And KeyAscii <= 122 Then
     Up = KeyAscii - 32
  Else
     Up = KeyAscii
  End If
    
Exit Function
 
ErrT93:
 
    RtCodeError = GestErr(Err, "Up")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Function ValData(Data As Control) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT94
 

   ValData = True
   
   If Len(Data) > 0 And Not IsDate(Data) Then
      MsgBox "Data specificata non valida (gg/mm/aa)"
      ValData = False
   End If

   

Exit Function
 
ErrT94:
 
    RtCodeError = GestErr(Err, "ValData")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Function VBstr(TheStr$) As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT95
 
' stripped einen Null terminerten String
' als VB string :
Dim TheTmp As String

     NullPos% = InStr(TheStr, Chr$(0))
     TheTmp = RTrim$(Left$(TheStr, NullPos% - 1))
     VBstr = TheTmp

Exit Function
 
ErrT95:
 
    RtCodeError = GestErr(Err, "VBstr")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function


Sub CopyFile(Source$, Dest$)
Dim RtCodeError As Integer
 
On Error GoTo ErrT96
 

Dim TheBuffer As String
Const BuffLen = 16384

    On Error GoTo errhandler


    Open Source$ For Binary Access Read As #1
    Open Dest$ For Binary Access Write As #2

    Buf% = 0
    If LOF(1) < BuffLen Then
       TheBuffer = Space$(LOF(1))
    Else
       TheBuffer = Space$(BuffLen)
    End If
    'MsgBox (Str$(Seek(1)) + " " + Str$(LOF(1)))

    Do While Seek(1) < LOF(1)

        'MsgBox (Str$(Seek(1)) + " " + Str$(LOF(1)))
        
        Buf% = Buf% + 1
        
        If LOF(1) - Seek(1) < BuffLen Then
           TheBuffer = Space$(LOF(1) - Seek(1) + 1)
           Get #1, , TheBuffer
           Put #2, , TheBuffer   ' Write to file.
           Exit Do
        Else
           Get #1, , TheBuffer
           Put #2, , TheBuffer   ' Write to file.
        End If
    
        'Call UpdateStatus(Len(TheBuffer), FALSE)

        i% = DoEvents()
    Loop

    Close #1
    Close #2
    Exit Sub
errhandler:
  '  warning (Error$(Err) & " -  " & Str(Err) & " - " & source$)
 '   Err_ Err, "errore durante la copia del file"
    Close #1
    Close #2
    Exit Sub
Exit Sub
 
ErrT96:
 
    RtCodeError = GestErr(Err, "CopyFile")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Function Doit(TheStr$) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT97
 
' default is YES
  i% = MsgBox(TheStr, 4 + 32, App.Title)
  If i% = 6 Then
    Doit = True
  Else
    Doit = False
  End If
Exit Function
 
ErrT97:
 
    RtCodeError = GestErr(Err, "Doit")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Sub Information(TheStr$)
Dim RtCodeError As Integer
 
On Error GoTo ErrT98
 
 i% = MsgBox(TheStr, 64, App.Title)
Exit Sub
 
ErrT98:
 
    RtCodeError = GestErr(Err, "Information")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Function IsValidPath%(ByVal DestPath$, ByVal DefaultDrive$)
Dim RtCodeError As Integer
 
On Error GoTo ErrT99
 


'   ======================================================
'   Remove left and right spaces
'   ======================================================
'    DestPath$ = AllTrim$(DestPath$)
'    DefaultDrive$ = AllTrim$(DefaultDrive$)

'   ======================================================
'   Check Default Drive Parameter
'   ======================================================
    If Right$(DefaultDrive$, 1) <> ":" Or Len(DefaultDrive$) <> 2 Then
        Msg$ = "Bad default drive parameter specified in IsValidPath "
        Msg$ = Msg$ + "Function.  You passed,  """ + DefaultDrive$ + """.  Must "
        Msg$ = Msg$ + "be one drive letter and "":"".  For "
        Msg$ = Msg$ + "example, ""C:"", ""D:""..."
        MsgBox Msg$, 64, "Setup Kit Error"
        GoTo parseErr
    End If


'   ======================================================
'   Insert default drive if path begins with root backslash
'   ======================================================
    If Left$(DestPath$, 1) = "\" Then
        DestPath$ = DefaultDrive + DestPath$
    End If


'   ======================================================
'   check for invalid characters
'   ======================================================
    On Error Resume Next
    Tmp$ = Dir$(DestPath$)
    If Err <> 0 Then
        GoTo parseErr
    End If


'   ======================================================
'   Check for wildcard characters and spaces
'   ======================================================
    If (InStr(DestPath$, "*") <> 0) Then GoTo parseErr
    If (InStr(DestPath$, "?") <> 0) Then GoTo parseErr
    If (InStr(DestPath$, " ") <> 0) Then GoTo parseErr


'   ======================================================
'   Make Sure colon is in second char position
'   ======================================================
    If Mid$(DestPath$, 2, 1) <> Chr$(58) Then GoTo parseErr


'   ======================================================
'   Insert root backslash if needed
'   ======================================================
    If Len(DestPath$) > 2 Then
      If Right$(Left$(DestPath$, 3), 1) <> "\" Then
        DestPath$ = Left$(DestPath$, 2) + "\" + Right$(DestPath$, Len(DestPath$) - 2)
      End If
    End If


'   ======================================================
'   Check drive to install on
'   ======================================================
    Drive$ = Left$(DestPath$, 1)
    ChDrive (Drive$)                        ' Try to change to the dest drive
    If Err <> 0 Then GoTo parseErr

'   ======================================================
'   Add final \
'   ======================================================
    If Right$(DestPath$, 1) <> "\" Then
        DestPath$ = DestPath$ + "\"
    End If


'   ======================================================
'   Root dir is a valid dir
'   ======================================================
    If Len(DestPath$) = 3 Then
        If Right$(DestPath$, 2) = ":\" Then
            GoTo ParseOK
        End If
    End If


'   ======================================================
'   Check for repeated Slash
'   ======================================================
    If InStr(DestPath$, "\\") <> 0 Then GoTo parseErr


'   ======================================================
'   Check for illegal directory names
'   ======================================================
    legalChar$ = "!#$%&'()-0123456789@ABCDEFGHIJKLMNOPQRSTUVWXYZ^_`{}~."
    BackPos = 3
    forePos = InStr(4, DestPath$, "\")
    Do
        temp$ = Mid$(DestPath$, BackPos + 1, forePos - BackPos - 1)

        '----------------------------
        'Test for illegal characters
        '----------------------------
        For i = 1 To Len(temp$)
            If InStr(legalChar$, UCase$(Mid$(temp$, i, 1))) = 0 Then GoTo parseErr
        Next i

        '-------------------------------------------
        'Check combinations of periods and lengths
        '-------------------------------------------
        periodPos = InStr(temp$, ".")
        Length = Len(temp$)
        If periodPos = 0 Then
            If Length > 8 Then GoTo parseErr                         'Base too long
        Else
            If periodPos > 9 Then GoTo parseErr                      'Base too long
            If Length > periodPos + 3 Then GoTo parseErr             'Extension too long
            If InStr(periodPos + 1, temp$, ".") <> 0 Then GoTo parseErr 'Two periods not allowed
        End If

        BackPos = forePos
        forePos = InStr(BackPos + 1, DestPath$, "\")
    Loop Until forePos = 0

ParseOK:
    IsValidPath% = True
    Exit Function

parseErr:
    IsValidPath% = False
Exit Function
 
ErrT99:
 
    RtCodeError = GestErr(Err, "IsValidPath%")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Sub RemoveFile(Fname$)
Dim RtCodeError As Integer
 
On Error GoTo ErrT100
 
If Doit("Delete " + Fname) = True Then
   Kill (Fname)
End If
Exit Sub
 
ErrT100:
 
    RtCodeError = GestErr(Err, "RemoveFile")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Function retry(TheStr$) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT101
 
  i% = MsgBox(TheStr, 5 + 32, App.Title)
  If i% = 4 Then
    retry = True
  Else
    retry = False
  End If

Exit Function
 
ErrT101:
 
    RtCodeError = GestErr(Err, "retry")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function


Sub AppRunning()
Dim RtCodeError As Integer
 
On Error GoTo ErrT102
 
        Dim sMsg As String
        If App.PrevInstance Then
        sMsg = App.EXEName & " Applicazione Già Attiva! "
           MsgBox sMsg, 4112
        End
        End If
Exit Sub
 
ErrT102:
 
    RtCodeError = GestErr(Err, "AppRunning")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Sub CenterForm(frmIn As Form)
Dim RtCodeError As Integer
 
On Error GoTo ErrT103
 
        Dim iTop As Integer, iLeft As Integer

        If frmIn.WindowState <> 0 Then Exit Sub
        iTop = (Screen.Height - frmIn.Height) \ 2
        iLeft = (Screen.Width - frmIn.Width) \ 2
    
        'If iTop And iLeft Then
        frmIn.Move iLeft, iTop
        'End If
Exit Sub
 
ErrT103:
 
    RtCodeError = GestErr(Err, "CenterForm")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Sub CenterMDIChild(frmParent As Form, frmChild As Form)
Dim RtCodeError As Integer
 
On Error GoTo ErrT104
 
        Dim iTop As Integer, iLeft As Integer
        If frmParent.WindowState <> 0 Or frmChild.WindowState <> 0 Then Exit Sub
        iTop = (frmParent.ScaleHeight - frmChild.Height) \ 2
        iLeft = (frmParent.ScaleWidth - frmChild.Width) \ 2

        If iTop And iLeft Then
        frmChild.Move iLeft, iTop
        End If
Exit Sub
 
ErrT104:
 
    RtCodeError = GestErr(Err, "CenterMDIChild")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub


Sub Critical(TheStr$)
Dim RtCodeError As Integer
 
On Error GoTo ErrT105
 
 i% = MsgBox(TheStr, 16 + 4096, App.Title) ' 16
Exit Sub
 
ErrT105:
 
    RtCodeError = GestErr(Err, "Critical")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Sub CutCopyPaste(iChoice As Integer)
Dim RtCodeError As Integer
 
On Error GoTo ErrT106
 
        ' ActiveForm refers to the active form in the MDI form.
        If TypeOf Screen.ActiveControl Is TextBox Then
        Select Case iChoice
                        Case 0          ' Cut.
                        ' Copy selected text to Clipboard.
                        Clipboard.SetText Screen.ActiveControl.SelText
                        ' Delete selected text.
                        Screen.ActiveControl.SelText = ""
                        Case 1          ' Copy.
                        ' Copy selected text to Clipboard.
                        Clipboard.SetText Screen.ActiveControl.SelText
                        Case 2          ' Paste.
                        ' Put Clipboard text in text box.
                        Screen.ActiveControl.SelText = Clipboard.GetText()
                        Case 3          ' Delete.
                        ' Delete selected text.
                        Screen.ActiveControl.SelText = ""
        End Select
        End If
Exit Sub
 
ErrT106:
 
    RtCodeError = GestErr(Err, "CutCopyPaste")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub


'*******************************************************
'* Procedure Name: FileExists                          *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function will check to make sure that a file    *
'*exists.It will return True if the file was found and *
'*False if it was not found.                           *
'*Example: If Not FileExists("autoexec.bat") Then...   *
'*******************************************************
Function FileExists(FileName As String) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT107
 
Dim i As Integer
On Error Resume Next
i = Len(Dir$(FileName))
If Err Or i = 0 Then
    FileExists = False
  Else
    FileExists = True
End If
Exit Function
 
ErrT107:
 
    RtCodeError = GestErr(Err, "FileExists")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

'*******************************************************
'* Procedure Name: FixAmper                            *
'*-----------------------------------------------------*
'* Created:           By: Bart Larsen                  *
'* Modified:          By:                              *
'*=====================================================*
'*This function fixes strings that are to be used by a *
'*Label control so the "&" (Chr(38)) does not          *
'*underscore the following character.                  *
'*[Label1.Caption=FixAmper("bf&ftr.zip")]              *
'*                                                     *
'*******************************************************
Function FixAmper(Strng As String) As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT108
 
Dim sTemp As String, N As Integer
    While InStr(Strng, "&")
        N = InStr(Strng, "&")
        sTemp = sTemp + Left$(Strng, N) + "&"
        Strng = Mid$(Strng, N + 1)
    Wend
    sTemp = sTemp + Strng
    FixAmper = sTemp
Exit Function
 
ErrT108:
 
    RtCodeError = GestErr(Err, "FixAmper")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

'x*******************************************************
'* Procedure Name: GetAppPath                          *
'*-----------------------------------------------------*
'* Created: 3/24/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*Returns the application path with a trailing \.      *
'*To use, call the function [SomeString=GetAppPath()]  *
'*                                                     *
'*                                                     *
'*                                                     *
'*******************************************************
Function GetAPPPath() As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT109
 
        Dim sTemp As String
        sTemp = App.Path
        If Right$(sTemp, 1) <> "\" Then sTemp = sTemp + "\"
        GetAPPPath = sTemp
Exit Function
 
ErrT109:
 
    RtCodeError = GestErr(Err, "GetAPPPath")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

Function GetUserName() As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT110
 
   
   
   

Exit Function
 
ErrT110:
 
    RtCodeError = GestErr(Err, "GetUserName")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

'*******************************************************
'* Procedure Name: CheckUnique                         *
'*-----------------------------------------------------*
'* Created: 4/18/94   By: KeepOnTop                    *
'* Modified:          By:                              *
'*=====================================================*
'*Keep form on top. Note that this is switched off if  *
'*form is minimised, so place in resize event as well. *
'*                                                     *
'*                                                     *
'*                                                     *
'*******************************************************
Sub KeepOnTop(f As Form)
Dim RtCodeError As Integer
 
On Error GoTo ErrT111
 
'Keep form on top. Note that this is switched off if form is
'minimised, so place in resize event as well.
' Const wFlags = SWP_NOMOVE Or SWP_NOSIZE

'    SetWindowPos f.hWnd, HWND_TOPMOST, 0, 0, 0, 0, wFlags    'Window will stay on top
    'To undo call again with HWND_NOTOPMOST
'    DoEvents
Exit Sub
 
ErrT111:
 
    RtCodeError = GestErr(Err, "KeepOnTop")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

'*******************************************************
'* Procedure Name: LongDirFix                          *
'*-----------------------------------------------------*
'* Created: 6/30/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function will shorten a directory name to the   *
'*length passed to it.                                 *
'*Usage: Label1.Caption=LongDirFix(sDirName, 32)       *
'*The second paramater the the max length of the       *
'*returned string.                                     *
'*******************************************************
Function LongDirFix(Incomming As String, Max As Integer) As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT112
 
Dim i As Integer, LblLen As Integer, StringLen As Integer
Dim TempString As String

TempString = Incomming
LblLen = Max

If Len(TempString) <= LblLen Then
    LongDirFix = TempString
    Exit Function
End If

LblLen = LblLen - 6

For i = Len(TempString) - LblLen To Len(TempString)
    If Mid$(TempString, i, 1) = "\" Then Exit For
        
Next

LongDirFix = Left$(TempString, 3) + "..." + Right$(TempString, Len(TempString) - (i - 1))

Exit Function
 
ErrT112:
 
    RtCodeError = GestErr(Err, "LongDirFix")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

'*******************************************************
'* Procedure Name: MakeDir                             *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function will create a directory even if the    *
'*underlying directories do not exist.                 *
'*Usage: MakeDir "c:\temp\demo"                        *
'*This procedue also uses the ValDir to find if the    *
'*directory already exists.                            *
'*******************************************************
Sub MakeDir(sDirName As String)
Dim RtCodeError As Integer
 
On Error GoTo ErrT113
 
Dim iMouseState As Integer
Dim iNewLen As Integer
Dim iDirLen As Integer

'Get Mouse State
iMouseState = Screen.MousePointer

'Change Mouse To Hour Glass
Screen.MousePointer = 11

'Set Start Length To Search For [\]
iNewLen = 4

'Add [\] To Directory Name If Not There
If Right$(sDirName, 1) <> "\" Then sDirName = sDirName + "\"

'Create Nested Directory
While Not ValDir(sDirName)
    iDirLen = InStr(iNewLen, sDirName, "\")
    If Not ValDir(Left$(sDirName, iDirLen)) Then
        MkDir Left$(sDirName, iDirLen - 1)
    End If
    iNewLen = iDirLen + 1
Wend

'Leave The Mouse The Way You Found It
Screen.MousePointer = iMouseState

Exit Sub
 
ErrT113:
 
    RtCodeError = GestErr(Err, "MakeDir")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

'*******************************************************
'* Procedure Name: ReadINI                             *
'*-----------------------------------------------------*
'* Created:           By: Daniel Bowen                 *
'* Modified: 3/24/94  By: David McCarter               *
'*=====================================================*
'*Returns a string from an INI file. To use, call the  *
'*functions and pass it the AppName, KeyName and INI   *
'*File Name, [sReg=ReadINI(App1,Key1,INIFile)]. If you *
'*need the returned value to be a integer then use the *
'*val command.                                         *
'*******************************************************
Function ReadIni(Appname As String, KeyName As String, FileName As String) As String
Dim RtCodeError As Integer
 
On Error GoTo ErrT114
 
Dim sRet As String
sRet = String(255, Chr(0))

ReadIni = Left(sRet, GetPrivateProfileString(Appname, ByVal KeyName, "", sRet, Len(sRet), FileName))

Exit Function
 
ErrT114:
 
    RtCodeError = GestErr(Err, "ReadIni")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

'*******************************************************
'* Procedure Name: SelectText                          *
'*-----------------------------------------------------*
'* Created: 2/14/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*Selects all the text in a text box. Call it when the *
'*text box get focus, [SelectText Text1.text]          *
'*                                                     *
'*                                                     *
'*                                                     *
'*******************************************************
Sub SelectText(ctrIn As Control)
Dim RtCodeError As Integer
 
On Error GoTo ErrT115
 
ctrIn.SelStart = 0
ctrIn.SelLength = Len(ctrIn.Text)
Exit Sub
 
ErrT115:
 
    RtCodeError = GestErr(Err, "SelectText")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub

Sub ShowAboutBox(frmIn As Form, sXtraInfo As String)
Dim RtCodeError As Integer
 
On Error GoTo ErrT116
 
  '  Call ShellAbout(frmIn.hWnd, App.Title, sXtraInfo, frmIn.Icon)
Exit Sub
 
ErrT116:
 
    RtCodeError = GestErr(Err, "ShowAboutBox")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Sub



'*******************************************************
'* Procedure Name: ValDir                              *
'*-----------------------------------------------------*
'* Created: 8/29/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*This function is used by MakeDir to validate if a    *
'*directory already exists.                            *
'*******************************************************
Function ValDir(sIncoming As String) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT117
 
Dim iCheck As String, iErrResult As Integer

On Local Error GoTo ValDirError

'If Right$(sIncoming, 1) <> "\" Then sIncoming = sIncoming + "\"
iCheck = Dir$(sIncoming)

If iErrResult = 76 Then
    ValDir = False
    Else
        ValDir = True
End If

Exit Function

ValDirError:

Select Case Err
    Case Is = 76
       iErrResult = Err
       Resume Next
    Case Else
End Select

Exit Function
 
ErrT117:
 
    RtCodeError = GestErr(Err, "ValDir")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

'*******************************************************
'* Procedure Name: WriteINI                            *
'*-----------------------------------------------------*
'* Created: 2/10/94   By: David McCarter               *
'* Modified:          By:                              *
'*=====================================================*
'*Writes a string to an INI file. To use, call the     *
'*function and pass it the AppName, KeyName, the New   *
'*String and the INI File Name,                        *
'*[R=WriteINI(App1,Key1,sReg,INIFile)]. Returns a 1 if *
'*there were no errors and a 0 if there were errors.   *
'*******************************************************
Function WriteINI(Appname As String, KeyName As String, NewString As Variant, FileName As String) As Integer
Dim RtCodeError As Integer
 
On Error GoTo ErrT118
 
     Dim s As Integer
     s = WritePrivateProfileString(Appname, KeyName, CStr(NewString), FileName)
     If s = False Then
         MsgBox "Problemi di scrittura configurazione. Impossibile continuare, chiamare l'assistenza tecnica per risolvere il problema!", 16, "STOP"
         End
     End If

Exit Function
 
ErrT118:
 
    RtCodeError = GestErr(Err, "WriteINI")
    Select Case RtCodeError
       Case vbAbort   ' Annulla
         End
       Case vbRetry   ' Riprova
         Resume
       Case vbIgnore   ' Ignora
         Resume Next
    End Select
End Function

