VERSION 5.00
Begin VB.Form wndAddFiles 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   4245
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7095
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4245
   ScaleWidth      =   7095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrVar 
      Interval        =   50
      Left            =   5580
      Top             =   3180
   End
   Begin VB.PictureBox pBar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00004000&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      ScaleHeight     =   225
      ScaleWidth      =   6585
      TabIndex        =   12
      Top             =   120
      Width           =   6615
      Begin VB.Label lblCaption 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Відкрити файли"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   6615
      End
   End
   Begin VB.CommandButton btnInternet 
      Caption         =   "Інтернет"
      Height          =   375
      Left            =   5580
      TabIndex        =   11
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton btnOthPath 
      Caption         =   "Стрибнути..."
      Height          =   375
      Left            =   5580
      TabIndex        =   10
      Top             =   1680
      Width           =   1215
   End
   Begin VB.DriveListBox lstDrive 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2760
      TabIndex        =   7
      Top             =   3600
      Width           =   2475
   End
   Begin VB.ComboBox lstHistory 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   660
      Width           =   2475
   End
   Begin VB.DirListBox lstDir 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   2340
      Left            =   2760
      TabIndex        =   5
      Top             =   1080
      Width           =   2475
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   4
      Top             =   660
      Width           =   2355
   End
   Begin VB.ComboBox lstTypes 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3600
      Width           =   2355
   End
   Begin VB.FileListBox lstFile 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FFFF&
      Height          =   2430
      Left            =   240
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   1080
      Width           =   2355
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Відмінити"
      Height          =   375
      Left            =   5580
      TabIndex        =   1
      Top             =   1140
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5580
      TabIndex        =   0
      Top             =   660
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00000000&
      Height          =   3855
      Left            =   120
      Top             =   240
      Width           =   6855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Папочки:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2760
      TabIndex        =   9
      Top             =   420
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Файли:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   420
      Width           =   1515
   End
End
Attribute VB_Name = "wndAddFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Integer, OldY As Integer



Private Sub btnInternet_Click()
Dim A, Z
For A = 1 To 1100: Z = Z + Chr(32 + (Rnd * (255 - 32))): Next A
MsgBox Z + " not available", vbInformation, "Internet not available!"
End Sub


Private Sub btnOthPath_Click()
Dim OPath As String
On Error Resume Next
OPath = InputBox("Вкажіть адесу тут:", "Go to folder", lstDir.Path)
If OPath = "" Then Exit Sub
lstDir.Path = OPath
lstDrive.Drive = Mid(OPath, 1, 3)
If Err Then MsgBox Err.Description, vbCritical, "ERROR " + Format(Err.Number)
End Sub

Private Sub Form_Activate()
If Me.Tag <> "Open" Then
 lblCaption.Caption = "Додати файл..."
Else
 lblCaption.Caption = "Відкрити файли..."
End If

Me.tmrVar.Enabled = True


End Sub

Private Sub Form_Deactivate()
  Me.tmrVar.Enabled = False
End Sub


Private Sub Form_Resize()

Dim X, Y, MX, MY, A, CL

MY = Me.Height / Screen.TwipsPerPixelY
For Y = 0 To MY
   Me.Line (0, Y * Screen.TwipsPerPixelY)-(Me.Width, Y * Screen.TwipsPerPixelY), RGB((100 / MY * Y), (100 / MY * Y), 0)
Next

' Dim a As Integer, CL As Integer
For A = 0 To Me.Height Step 25
  CL = 70 + (40 * Sin(A))
  Me.Line (0, A)-(Me.Width, A), RGB(CL, CL, 0), BF
Next

End Sub



Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X
OldY = Y
End Sub


Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ML, MT

If Button = 1 Then

 ML = Me.Left + (X - OldX)
 MT = Me.Top + (Y - OldY)
 
 If ML <= 120 Then ML = 0
 If MT <= 120 Then MT = 0
 
 If ML > Screen.Width - Me.Width - 120 Then ML = Screen.Width - Me.Width
 If MT > Screen.Height - Me.Height - 120 Then MT = Screen.Height - Me.Height
 
 Me.Left = ML
 Me.Top = MT

 X = ML
 Y = MT
 
End If

End Sub



Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub Form_Load()

  lstTypes.AddItem "All Supported"
  lstTypes.AddItem "Wave Files (*.wav)"
  lstTypes.AddItem "Midi Files (*.mid)"
  lstTypes.AddItem "Riff Midi Files (*.rmi)"
  lstTypes.AddItem "Windows Media Audio Files (*.wma)"
  lstTypes.AddItem "MXT Playlists"
  lstTypes.AddItem "All Files (*.*)"
  
  
  lstTypes.ListIndex = 0
  
  
End Sub

Private Sub Form_Paint()
On Local Error Resume Next

Dim F, lPath$
F = FreeFile

Open LowPath(App.Path) + "history.mb1" For Input As #F
If Err Then Close: Exit Sub
lstHistory.Clear
Do While Not EOF(F)
  Line Input #F, lPath$
  lstHistory.AddItem lPath$
Loop
Close F

End Sub

Private Sub lstDir_Change()
lstFile.Path = lstDir.Path
End Sub

Private Sub lstDrive_Change()

On Error Resume Next

lstDir.Path = lstDrive.Drive

If Err Then
 Err.Clear
 lstDrive.Drive = Left(lstDir.Path, 2)
 MsgBox "Диск ''" + lstDrive.Drive + "'' не реагує!", vbExclamation
End If

End Sub


Private Sub lstFile_DblClick()
OKButton_Click
End Sub

Private Sub lstFile_PatternChange()
 txtFileName.Text = lstFile.Pattern
End Sub





Private Sub lstHistory_Click()
 On Error Resume Next
 lstDir.Path = lstHistory.List(lstHistory.ListIndex)
 lstDrive.Drive = Left(lstHistory.List(lstHistory.ListIndex), 2)
End Sub

Private Sub lstTypes_Click()
If lstTypes.ListIndex = 0 Then lstFile.Pattern = "*.wav;*.mid;*.rmi;*.wma;*.mp?;*.avi;*.mxt"
If lstTypes.ListIndex = 1 Then lstFile.Pattern = "*.wav"
If lstTypes.ListIndex = 2 Then lstFile.Pattern = "*.mid"
If lstTypes.ListIndex = 3 Then lstFile.Pattern = "*.rmi"
If lstTypes.ListIndex = 4 Then lstFile.Pattern = "*.wma"
If lstTypes.ListIndex = 5 Then lstFile.Pattern = "*.mxt"
If lstTypes.ListIndex = 6 Then lstFile.Pattern = "*.*"

End Sub
Private Sub OKButton_Click()

Dim X, Y, TT As Integer


If Me.Tag <> "Open" Then Y = wndMain.lstSec.ListIndex
If Me.Tag = "Open" Then wndMain.lstMain.Clear: wndMain.lstSec.Clear: wndMain.lstTimes.Clear

For X = 0 To lstFile.ListCount - 1
 If lstFile.Selected(X) = True Then
  If UCase(Right(lstFile.List(X), 3)) = "MXT" Then
    wndMain.LoadListAdd LowPath(lstFile.Path) + lstFile.List(X), wndMain.lstMain, wndMain
  Else
    wndMain.lstMain.AddItem LowPath(lstFile.Path) + lstFile.List(X)
    wndMain.lstSec.AddItem GetMp3Song(LowPath(lstFile.Path) + lstFile.List(X), TT)
    wndMain.lstTimes.AddItem Format(TT)
  End If
 End If
Next


If Me.Tag <> "Open" Then wndMain.lstSec.ListIndex = Y

If Me.Tag = "Open" Then
   If wndMain.lstMain.ListCount > 0 Then
     wndMain.lstSec.ListIndex = 0
     Call wndMain.ResetStatus
     Call wndMain.Command_Open
     Call wndMain.Command_Play
   End If
End If

HistoryUpdate


Me.Hide

End Sub

Private Sub tmrVar_Timer()
  
  
    
  Dim X
  
  X = 50 + (50 * Sin(Timer * 1))
  
  pBar.Scale (X, 0)-(X + 100, 1)
  
  Dim QW As Integer
  
  For QW = 0 To 100
   pBar.Line (QW, 0)-(QW + 1, 1), RGB(155 / 255 * QW, 155 / 255 * QW, 0), BF
  Next
  
  For QW = 0 To 100
   pBar.Line (100 + QW, 0)-(100 + QW + 1, 1), RGB(155 / 255 * (100 - QW), 155 / 255 * (100 - QW), 0), BF
  Next

End Sub

Private Sub txtFileName_Change()
 On Error Resume Next
 lstFile.Pattern = txtFileName.Text
 If Err Then MsgBox "Unsupported symbols found!", vbCritical
End Sub

Sub HistoryUpdate()

On Error Resume Next
Dim W, File_Z$, Present, Modo
W = FreeFile

Open LowPath(App.Path) + "history.mb1" For Input As #W
   Present = 0
   Do
   Input #W, File_Z$
   If UCase$(File_Z$) = UCase$(lstDir.Path) Then Present = 1
   Loop While Not EOF(W)
Close W
   
If Present = 1 Then Exit Sub

Open LowPath(App.Path) + "history.mb1" For Append As #W
   Print #W, lstDir.Path
Close W
   
End Sub


