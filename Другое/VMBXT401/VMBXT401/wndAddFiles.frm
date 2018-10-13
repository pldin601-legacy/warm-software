VERSION 5.00
Begin VB.Form wndAddFiles 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Добавить файлы в список..."
   ClientHeight    =   3735
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   7020
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "wndAddFiles.frx":0000
   ScaleHeight     =   3735
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton INet 
      Caption         =   "Download"
      Enabled         =   0   'False
      Height          =   375
      Left            =   5640
      TabIndex        =   10
      Top             =   1080
      Width           =   1215
   End
   Begin VB.DriveListBox lstDrive 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2700
      TabIndex        =   8
      Top             =   3300
      Width           =   2475
   End
   Begin VB.ComboBox lstHistory 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   2700
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2475
   End
   Begin VB.DirListBox lstDir 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2340
      Left            =   2700
      TabIndex        =   5
      Top             =   780
      Width           =   2475
   End
   Begin VB.TextBox txtFileName 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      TabIndex        =   4
      Top             =   360
      Width           =   2355
   End
   Begin VB.ComboBox lstTypes 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   3300
      Width           =   2355
   End
   Begin VB.FileListBox lstFile 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2430
      Left            =   180
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   780
      Width           =   2355
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Папки:"
      Height          =   195
      Left            =   2700
      TabIndex        =   9
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Файлы:"
      Height          =   195
      Left            =   180
      TabIndex        =   6
      Top             =   120
      Width           =   1515
   End
End
Attribute VB_Name = "wndAddFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Me.Hide
End Sub

Private Sub Form_Load()

  lstTypes.AddItem "Wave Files (*.wav)"
  lstTypes.AddItem "Midi Files (*.mid)"
  lstTypes.AddItem "Riff Midi Files (*.rmi)"
  lstTypes.AddItem "Windows Media Audio Files (*.wma)"
  lstTypes.AddItem "All Other Formats"
  lstTypes.AddItem "All Files (*.*)"
  
  
  lstTypes.ListIndex = 0
  
  Label1.Caption = Language("FILES+:")
  Label2.Caption = Language("DIRS+:")
  Me.Caption = Language("FORM_ADD:")

End Sub

Private Sub Form_Paint()
On Local Error Resume Next

Dim F, lPath$
F = FreeFile

Open LowPath(App.path) + "history.mb1" For Input As #F
If Err Then Close: Exit Sub

lstHistory.Clear
Do While Not EOF(F)
  Line Input #F, lPath$
  lstHistory.AddItem lPath$
Loop

Close F
End Sub

Private Sub lstDir_Change()
lstFile.path = lstDir.path
End Sub

Private Sub lstDrive_Change()
On Error Resume Next
lstDir.path = lstDrive.Drive

If Err Then
 MsgBox "Drive not ready!", vbExclamation
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
 lstDir.path = lstHistory.List(lstHistory.ListIndex)
 lstDrive.Drive = Left(lstHistory.List(lstHistory.ListIndex), 2)
End Sub

Private Sub lstTypes_Click()
If lstTypes.ListIndex = 0 Then lstFile.Pattern = "*.wav"
If lstTypes.ListIndex = 1 Then lstFile.Pattern = "*.mid"
If lstTypes.ListIndex = 2 Then lstFile.Pattern = "*.rmi"
If lstTypes.ListIndex = 3 Then lstFile.Pattern = "*.wma"
If lstTypes.ListIndex = 4 Then lstFile.Pattern = "*.mpg;*.dat;*.mp3;*.mp2;*.avi"
If lstTypes.ListIndex = 5 Then lstFile.Pattern = "*.*"

End Sub

Private Sub OKButton_Click()

  Dim x
  
For x = 0 To lstFile.ListCount - 1
 If lstFile.Selected(x) = True Then
  wndMain.lstMain.AddItem LowPath(lstFile.path) + lstFile.List(x)
  wndMain.lstMain.Selected(wndMain.lstMain.ListCount - 1) = True
 End If
Next

HistoryUpdate
Me.Hide

End Sub

Private Sub txtFileName_Change()
 lstFile.Pattern = txtFileName.Text
End Sub

Sub HistoryUpdate()

On Error Resume Next
Dim W, File_Z$, Present, Modo
W = FreeFile

Open LowPath(App.path) + "history.mb1" For Input As #W
   Present = 0
   Do
   Input #W, File_Z$
   If UCase$(File_Z$) = UCase$(lstDir.path) Then Present = 1
   Loop While Not EOF(1)
Close W
   
If Present = 1 Then Exit Sub

Open LowPath(App.path) + "history.mb1" For Append As #W
   Print #W, lstDir.path
Close W
   
End Sub
