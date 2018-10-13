Attribute VB_Name = "VMB_Subs"
Type RecentFileHdr
  rfName As String * 256
  rfDescription As String * 256
  rfUsed As Boolean
End Type

Global BackIndex As Integer
Global RecFileHead As RecentFileHdr




Sub GetScreenResolution(rWidth As Integer, rHeight As Integer)
 Dim H_RES As Integer
 Dim P_RES As Integer
 
 H_RES = Screen.Height / Screen.TwipsPerPixelY
 
 Select Case H_RES
 
  Case 480: rWidth = 640: rHeight = 480
  Case 576: rWidth = 720: rHeight = 576
  Case 600: rWidth = 800: rHeight = 600
  Case 768: rWidth = 1024: rHeight = 768
  Case 864: rWidth = 1152: rHeight = 864
  
 End Select
 
End Sub


Function GetTimeFromMinutes(vMinutes As Integer)
 GetTimeFromMinutes = Format$(Fix(vMinutes / 60), "00") & ":" & Format$(Fix(vMinutes Mod 60), "00")
End Function

Function GetUserMediaId(Filename As String) As String
' Constructing
GetUserMediaId = ""
End Function




Function InStrX(A, B, C) As Long
  For Z = Len(B) To 1 Step -1
    If InStr(Z, B, C) > 0 Then
      InStrX = Z
      Exit For
    End If
  Next
End Function


Function Language(Rec As String)

Language = GetIniRecord(Rec, LowPath(App.Path) + GetIniRecord("LANG_FILE:", LowPath(App.Path) + "LANGUAGE.SEL"))

End Function

Sub SaveFileHdr(mxFile As String, mxName As String)

On Error Resume Next

Dim I, J, K

I = FreeFile

Open (LowPath(App.Path) + "favourites.vmb") For Random As #I Len = Len(RecFileHead)

J = LOF(I) / Len(RecFileHead)

For K = 1 To J + 1
  
  Get #I, K, RecFileHead
  
  If RecFileHead.rfUsed = False Then
    RecFileHead.rfName = mxFile
    RecFileHead.rfDescription = mxName
    RecFileHead.rfUsed = True
    Put #I, K, RecFileHead
    Exit For
  End If

Next

Close I

End Sub

Sub RemoveFileHdr(mxName As String)

On Error Resume Next

Dim I, J, K

I = FreeFile

Dim RecFileHead As RecentFileHdr

Open (LowPath(App.Path) + "favourites.vmb") For Random As #I Len = Len(RecFileHead)

J = LOF(I) / Len(RecFileHead)

For K = 1 To J
  Get #I, K, RecFileHead
  If RecFileHead.rfUsed = True Then
   If Trim(RecFileHead.rfDescription) = mxName Then
    RecFileHead.rfUsed = False
    Put #I, K, RecFileHead
    Exit For
   End If
  End If
Next

Close I

End Sub


Sub Sleep(tm As Currency)
Dim Tx As Currency
Tx = Timer
Do: Loop While Not Timer >= Tx + tm
End Sub

Sub SleepVD(tm As Currency)
Dim Tx As Currency
Tx = Timer
Do: wndMain.tmPos.TimeSet = "00:" + Format(Timer - (Tx + tm), "00"): DoEvents: Loop While Not Timer >= Tx + tm
End Sub

Sub SleepX(tm As Currency)
Dim Tx As Currency
Tx = Timer
Do: DoEvents: Loop While Not Timer >= Tx + tm
End Sub

Function WINTODOS(Tex$) As String
Dim Char, Cit, Zen As String
Zen = Tex$

For X = 1 To Len(Tex$)

Char = Asc(Mid$(Zen, X, 1))
If Char >= 192 And Char <= 239 Then
  Cit = Char - 64
  GoTo OK
Else
  Cit = Char
End If

If Char >= 240 And Char <= 255 Then
  Cit = Char - 16
  GoTo OK
Else
  Cit = Char
End If

OK:
Mid$(Zen, X, 1) = Chr$(Cit)

Next

WINTODOS = Zen

End Function


Public Function GetVersion() As String
GetVersion = Format$(App.Major, "0") + "." + Format$(App.Minor, "00")
End Function
Function PathHead$(Filename As String)
Dim Names As Integer
For Names = Len(Filename) To 1 Step -1
 If Mid$(Filename, Names, 1) = "\" Then
  PathHead$ = Mid$(Filename, 1, (Names) - 1)
  If PathHead$ = "$APPDIR$" Then PathHead$ = App.Path
  Exit For
 End If
Next
End Function
Function FileExists(Path$) As Boolean

    On Error Resume Next
    
    Dim X As Integer

    X = FreeFile

    On Error Resume Next
    Open Path$ For Input As X
    If Err = 0 Then
        FileExists = True
    Else
        FileExists = False
    End If
    Close X

End Function
Public Function FileHead$(Filename As String)
Dim Names As Integer
For Names = Len(Filename) To 1 Step -1
If Mid$(Filename, Names, 1) = "\" Then FileHead$ = Right$(Filename, Len(Filename) - (Names)): Exit Function
Next
End Function

Public Function LowPath(InPath As String) As String
If Right$(InPath, 1) = "\" Then LowPath = InPath
If Right$(InPath, 1) <> "\" Then LowPath = InPath + "\"
End Function

Public Function GetIniRecord(Record As String, INIFile As String)
Dim CfgLine As String, G As Integer
On Error Resume Next
G = FreeFile
Open INIFile For Input As #G
Do
Line Input #G, CfgLine
If UCase$(Mid$(CfgLine, 1, Len(Record))) = UCase(Record) Then
   GetIniRecord = Mid$(CfgLine, Len(Record) + 1)
End If
Loop While Not EOF(G)
Close G
End Function

Public Function GetMp3Song(ByRef Song As String, TimeOut As Integer)

' On Error Resume Next

If GlbNoTags Then GetMp3Song = Song: Exit Function

Dim xTITLE As String, xALBUM As String, xARTIST As String
Dim xYEAR As String, xGERNE As Integer, xCOMMENT As String
Dim xFILENAME As String, xFILEPATH As String, xFILETYPE As String
Dim xTIP As String
Dim CParm As String

If Mid(Song, 1, 2) = "*C" Then
  If UCase(Mid(Song, 3, 1)) = "J" Then
    CParm = Mid(Song, 4)
    Song = "Jump to " + CParm
  End If
End If

If FileExists(Song) = False Then
 GetMp3Song = Song
 Exit Function
End If

wndMain.lblMask.Caption = wndMain.lblMask.Caption + vbCrLf + Song

xTIP = Right(UCase(Song), Len(Song) - InStrX(1, Song, "."))

If xTIP = "MP3" Or Tip = "WMA" Then
 GetMP3idTAG Song, xTITLE, xARTIST, xALBUM, xYEAR, xCOMMENT, xGERNE
End If



If xARTIST = "" Then xARTIST = "Unknown Artist"
If xALBUM = "" Then xALBUM = "Unknown Album"
If xYEAR = "" Then xYEAR = "1900"

If xTIP = "MID" Then
 xTITLE = MIDI_NAME(Song)
End If

' If xTITLE = "" Then xTITLE = Mid(FileHead(Song), 1, Len(FileHead(Song)) - (Len(Song) - (InStr(1, Song, ".")) + 1))

xFILETYPE = xTIP
Close #I
    
    Dim PROT, PROTM

    wndMain.TimeX.Command = "close"
    wndMain.TimeX.Filename = Song
    wndMain.TimeX.Command = "open"
    wndMain.TimeX.TimeFormat = 0
    PROT = Fix(wndMain.TimeX.Length / 1000)
    PROTM = Format(Fix(PROT / 60), "00") + ":" + Format(PROT Mod 60, "00")
    wndMain.TimeX.Filename = ""
    wndMain.TimeX.Command = "close"
    TimeOut = CInt(PROT)


If xTITLE > "" Then wndMain.inTitle.Text = GlbTitle Else wndMain.inTitle.Text = GlbTagLess

If InStr(1, wndMain.inTitle.Text, "%11") > 0 Then
 
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%11") - 1
   wndMain.inTitle.SelLength = 3
   wndMain.inTitle.SelText = PROTM
 Loop While Not InStr(1, wndMain.inTitle.Text, "%11") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%10") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%10") - 1
   wndMain.inTitle.SelLength = 3
   wndMain.inTitle.SelText = Midi_Size(Song)
 Loop While Not InStr(1, wndMain.inTitle.Text, "%10") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%TAB") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%TAB") - 1
   wndMain.inTitle.SelLength = 4
   wndMain.inTitle.SelText = Chr(vbKeyTab)
 Loop While Not InStr(1, wndMain.inTitle.Text, "%TAB") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%1") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%1") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = Mid(FileHead(Song), 1, Len(FileHead(Song)) - (Len(Song) - (InStrX(1, Song, ".")) + 1))
 Loop While Not InStr(1, wndMain.inTitle.Text, "%1") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%2") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%2") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = LowPath(PathHead(Song))
 Loop While Not InStr(1, wndMain.inTitle.Text, "%2") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%3") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%3") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = xTIP
 Loop While Not InStr(1, wndMain.inTitle.Text, "%3") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%4") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%4") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = xTITLE
 Loop While Not InStr(1, wndMain.inTitle.Text, "%4") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%5") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%5") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = xARTIST
 Loop While Not InStr(1, wndMain.inTitle.Text, "%5") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%6") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%6") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = xALBUM
 Loop While Not InStr(1, wndMain.inTitle.Text, "%6") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%7") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%7") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = xYEAR
 Loop While Not InStr(1, wndMain.inTitle.Text, "%7") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%8") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%8") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = GetGerneString(xGERNE)
 Loop While Not InStr(1, wndMain.inTitle.Text, "%8") = 0
End If

If InStr(1, wndMain.inTitle.Text, "%9") > 0 Then
 Do
   wndMain.inTitle.SelStart = InStr(1, wndMain.inTitle.Text, "%9") - 1
   wndMain.inTitle.SelLength = 2
   wndMain.inTitle.SelText = xTIP
 Loop While Not InStr(1, wndMain.inTitle.Text, "%9") = 0
End If

GetMp3Song = wndMain.inTitle.Text

End Function


Public Function GetHTTPSong(ByRef Song As String)

On Error Resume Next

Dim RU As String
Dim TT As String, AU As String, AL As String
Dim YR As String, CM As String, Tip As String, GR As Integer

Tip = Right(UCase(Song), 3)

If Tip = "MP3" Or Tip = "MP2" Or Tip = "MP1" Or Tip = "WMA" Then
 GetMP3idTAG Song, TT, AU, AL, YR, CM, GR
End If

If Tip = "MID" Or Tip = "RMI" Then
 TT = MIDI_NAME(Song)
 YR = Midi_Size(Song)
End If

If TT = "" Then TT = Mid(FileHead(Song), 1, Len(FileHead(Song)) - 4)
If AU > "" Then AU = AU + " - "
If YR > "" Then YR = " (" + YR + ")"


If Tip = "WAV" Then RU = "[WAVE] "
If Tip = "MID" Then RU = "[MIDI] "
If Tip = "RMI" Then RU = "[MIDI] "
If Tip = "MP3" Then RU = "[MPEG] "
If Tip = "MP2" Then RU = "[MPEG] "
If Tip = "MP1" Then RU = "[MPEG] "
If Tip = "MPG" Then RU = "[VIDEO] "
If Tip = "MPG" Then RU = "[VDEO] "
If Tip = "WMA" Then RU = "[WMED] "

If RU = "" Then RU = "[????] "

GetHTTPSong = AU + TT

End Function


Function ReadCommand(ByRef GetCommand As String, ByRef GetValue As Boolean)
 If GetValue = True Then ReadCommand = Right$(GetCommand, Len(GetCommand) - 12)
 If GetValue = False Then ReadCommand = Mid$(GetCommand, 1, 11)
End Function

Function FilterName(Text As String) As String

Dim Ls, Bs, Variants, Bizer
On Error Resume Next

For Ls = 1 To Len(Text)
Bs = Mid$(Text, Ls, 1)

 For Variants = 0 To 47
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 91 To 96
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 58 To 63
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 For Variants = 123 To 191
  If Bs = Chr$(Variants) Then Bs = "_"
 Next
 
Mid$(Text, Ls, 1) = Bs

Next

If Text = "" Then Text = "Unnamed"
FilterName = Text

End Function

Sub ExchangeFiles(SI As Integer, DI As Integer, Sources As ListBox)
On Error Resume Next
Dim A, B, ASel As Boolean, BSel As Boolean
A = Sources.List(DI)
B = Sources.List(SI)
ASel = Sources.Selected(DI)
BSel = Sources.Selected(SI)
Sources.List(DI) = B
Sources.List(SI) = A
Sources.Selected(DI) = BSel
Sources.Selected(SI) = ASel

End Sub


