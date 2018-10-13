VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Stinger 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Quadrosoft MacroSound"
   ClientHeight    =   3285
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   6015
   Icon            =   "stinger.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   6015
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Com1 
      Left            =   4380
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Wave | *.wav"
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   35
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   2880
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   34
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   2880
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   33
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   2880
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   32
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   2880
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   31
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   2520
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   30
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   2520
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   29
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   2520
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   28
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   2520
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   27
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   26
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   25
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   24
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   2160
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   23
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   1800
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   22
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   1800
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   21
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   1800
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   20
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1800
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   19
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   18
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   17
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   16
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1440
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   15
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   14
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   13
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   12
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1080
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   11
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   10
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   9
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   8
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   720
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   7
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   6
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   5
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   360
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   3
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   0
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   2
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   0
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Height          =   375
      Index           =   1
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   0
      Width           =   1515
   End
   Begin VB.CommandButton StingBut 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   0
      Width           =   1515
   End
   Begin VB.Menu mmFile 
      Caption         =   "File"
      Begin VB.Menu mExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu nConfig 
      Caption         =   "Configure"
      Begin VB.Menu mfff 
         Caption         =   "Find File"
         Shortcut        =   ^I
      End
      Begin VB.Menu mmdelf 
         Caption         =   "Delete File"
         Shortcut        =   ^J
      End
      Begin VB.Menu t 
         Caption         =   "-"
      End
      Begin VB.Menu mmCC 
         Caption         =   "Change color"
         Shortcut        =   ^D
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu sstopp 
         Caption         =   "Stop playing"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "Stinger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_dx As New DirectX7
Dim m_ds As DirectSound
Dim m_dsBuffer(0) As DirectSoundBuffer

Dim INXS As Integer

Private Sub Form_Load()
Dim RCont As Integer, I As Integer
On Error Resume Next
Dim ST_MOVE As ST_SAVE
Open LowPath(App.Path) + "vmbxstg.cfg" For Random As #1 Len = Len(ST_MOVE)
If Err Then Exit Sub

RCont = LOF(1) / Len(ST_MOVE)

For I = 1 To RCont

Get #1, I, ST_MOVE
StingBut(I - 1).ToolTipText = Trim(ST_MOVE.st_FileInfo)
StingBut(I - 1).Caption = Trim(ST_MOVE.st_FileInfo)
StingBut(I - 1).Tag = Trim(ST_MOVE.st_FileName)
StingBut(I - 1).BackColor = Trim(ST_MOVE.st_Color)

Next I

Close #1


Err.Clear


End Sub

Private Sub Form_Unload(Cancel As Integer)

Dim I As Integer
On Error Resume Next

Dim ST_MOVE As ST_SAVE
Open LowPath(App.Path) + "vmbxstg.cfg" For Random As #1 Len = Len(ST_MOVE)
If Err Then MsgBox "Config Saveing Error.", vbCritical
If Err Then Exit Sub


For I = 1 To StingBut.Count

ST_MOVE.st_FileInfo = StingBut(I - 1).ToolTipText
ST_MOVE.st_FileInfo = StingBut(I - 1).Caption
ST_MOVE.st_FileName = StingBut(I - 1).Tag
ST_MOVE.st_Color = StingBut(I - 1).BackColor
Put #1, I, ST_MOVE

Next I

Close #1


End Sub


Private Sub mExit_Click()
Unload Stinger
End Sub

Private Sub mfff_Click()
On Error Resume Next

Dim Info As String
Com1.ShowOpen
If Err Then Exit Sub
StingBut(INXS).Tag = Com1.Filename
Info = InputBox("Enter the file description.", , FileHead(Com1.Filename))

If Info = "" Then Info = FileHead(Com1.Filename)
StingBut(INXS).Caption = Info
StingBut(INXS).ToolTipText = Info

End Sub

Private Sub mmCC_Click()
On Error Resume Next
Com1.ShowColor
If Err Then Exit Sub
StingBut(INXS).BackColor = Com1.Color
End Sub

Private Sub mmdelf_Click()
On Error Resume Next
Dim Vaga
Vaga = MsgBox("Delete file from this button?", vbQuestion + vbYesNo)
If Vaga = 7 Then Exit Sub
StingBut(INXS).Caption = ""
StingBut(INXS).ToolTipText = ""
StingBut(INXS).Tag = ""

End Sub


Private Sub sstopp_Click()
If m_dsBuffer(0) Is Nothing Then Exit Sub
m_dsBuffer.Stop
m_dsBuffer.SetCurrentPosition 0
End Sub

Private Sub StingBut_Click(Index As Integer)

On Error Resume Next


Set m_ds = m_dx.DirectSoundCreate("")
If Err Then MsgBox "Unable to initialize DirectSound!", vbCritical, "DirectX 7.0": End

m_ds.SetCooperativeLevel Me.hWnd, DSSCL_NORMAL

Dim bufferDesc As DSBUFFERDESC 'этот новый объект передаетс€ DS, чтобы описать, какой буфер создаетс€
'ќчень похоже на описание поверхности DirectDraw
Dim waveFormat As WAVEFORMATEX
'Ёти утановки выполн€ютс€ почти дл€ каждого приложени€
bufferDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC

waveFormat.nFormatTag = WAVE_FORMAT_PCM
waveFormat.nChannels = 2    'ƒва канала
waveFormat.lSamplesPerSec = 44100 '22 kHz можете конечно мен€ть, если разбираетесь.......
waveFormat.nBitsPerSample = 16  '16 бит лучше чем 8 (лучше качество)
waveFormat.nBlockAlign = waveFormat.nBitsPerSample / 8 * waveFormat.nChannels
waveFormat.lAvgBytesPerSec = waveFormat.lSamplesPerSec * waveFormat.nBlockAlign


Set m_dsBuffer(0) = m_ds.CreateSoundBufferFromFile(StingBut(INXS).Tag, bufferDesc, waveFormat)

m_dsBuffer(0).SetPan 0
m_dsBuffer(0).SetVolume 0

m_dsBuffer(0).Play 0
If Err Then MsgBox "DS ERROR PLAY", vbCritical
End Sub

Private Sub StingBut_GotFocus(Index As Integer)
INXS = Index

End Sub


Private Sub StingBut_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
StingBut(Index).SetFocus
If Button = 2 Then PopupMenu nConfig
End Sub


