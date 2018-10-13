VERSION 5.00
Begin VB.Form infForm 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   4650
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnOK 
      Cancel          =   -1  'True
      Caption         =   "Escape"
      Height          =   315
      Left            =   3420
      TabIndex        =   2
      Top             =   4080
      Width           =   1155
   End
   Begin VB.PictureBox panZone 
      BackColor       =   &H00000000&
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   4575
      TabIndex        =   1
      Top             =   360
      Width           =   4635
      Begin VB.CommandButton btnChange 
         Caption         =   "Change"
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   915
      End
      Begin VB.Label oAddress 
         Caption         =   "0000H-FFFFH"
         Height          =   195
         Left            =   1140
         TabIndex        =   10
         Top             =   1320
         Width           =   3375
      End
      Begin VB.Label iAddress 
         Caption         =   "Адрес:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label oDevice 
         Caption         =   "Set as default. Type default."
         Height          =   195
         Left            =   1140
         TabIndex        =   8
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label oFileSize 
         Caption         =   "0 Bytes (0Kb)"
         Height          =   195
         Left            =   1140
         TabIndex        =   7
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label oFileName 
         Caption         =   "..."
         Height          =   735
         Left            =   1140
         TabIndex        =   6
         Top             =   60
         Width           =   3375
      End
      Begin VB.Label iDevice 
         Caption         =   "Девайс:"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   1080
         Width           =   1035
      End
      Begin VB.Label iFileSize 
         Caption         =   "Розмір:"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   1035
      End
      Begin VB.Label iFileName 
         Caption         =   "Файл:"
         Height          =   735
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   1035
      End
   End
   Begin VB.Label lblCaption 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Інформація"
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
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4635
   End
End
Attribute VB_Name = "infForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX As Integer, OldY As Integer

Sub RefreshThis()
On Error Resume Next
Dim A As String, X, B, mBit As String * 1024, xBit As String * 1, zBit As String

Open oFileName.Caption For Input As #7
Me.Tag = wndMain.lstSec.ListIndex
If Err Then
 oFileSize = "File not found"
 Close #7
 Exit Sub
End If

If LOF(7) >= 1000 And LOF(7) < 1000000 Then oFileSize = Format(LOF(7), "### ### ### ##0") + " bytes (" + Format(LOF(7) / CCur(1000), "#.#") + "K)"
If LOF(7) >= 1000000 And LOF(7) < 1000000000 Then oFileSize = Format(LOF(7), "### ### ### ##0") + " bytes (" + Format(LOF(7) / CCur(1000000), "#.0") + "M)"
If LOF(7) >= 1000000000 Then oFileSize = Format(LOF(7), "### ### ### ##0") + " bytes (" + Format(LOF(7) / CCur(1000) / CCur(1000) / CCur(1000), "#.0") + "G)"
A = Hex(LOF(7))
oAddress.Caption = String(Len(A), "0") + "H-" + A + "H"
Close #7

End Sub

Private Sub btnChange_Click()
On Error Resume Next
Dim CNM As String, K, NN As Integer
CNM = InputBox("Змінити ім'я та адресу файла" + vbCrLf + oFileName.Caption + vbCrLf + "to", "Змінити ім'я", oFileName.Caption)
If CNM > "" Then
 If FileExists(CNM) = False Then
   K = MsgBox("File ''" + CNM + "'' не є існуючим. Ви дійсно хочете зробити цей крок?", vbExclamation + vbYesNo, "Warning!")
   If K = vbNo Then Exit Sub
   If K = vbYes Then
    wndMain.lstMain.List(Val(Me.Tag)) = CNM
    wndMain.lstSec.List(Val(Me.Tag)) = GetMp3Song(CNM, NN)
    wndMain.lstTimes.List(Val(Me.Tag)) = Str(NN)
    oFileName.Caption = CNM: RefreshThis
   End If
 Else
  wndMain.lstMain.List(Me.Tag) = CNM
  wndMain.lstSec.List(Me.Tag) = GetMp3Song(CNM, NN)
  wndMain.lstTimes.List(Val(Me.Tag)) = Str(NN)
  oFileName.Caption = CNM: RefreshThis
 End If
End If
End Sub

Private Sub btnOk_Click()
Me.Hide
End Sub


Private Sub Form_Activate()
RefreshThis

End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X
OldY = Y
End Sub


Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
 Me.Move Me.Left + (X - OldX), Me.Top + (Y - OldY)
End If

End Sub


