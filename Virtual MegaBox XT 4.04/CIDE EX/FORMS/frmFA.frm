VERSION 5.00
Begin VB.Form frmFA 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quadrosof Virtual MegaBox XT 4.01 File Associator"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5790
   Icon            =   "frmFA.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5790
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   2115
      Left            =   180
      TabIndex        =   2
      Top             =   180
      Width           =   5475
      Begin VB.CheckBox Wav_Mid 
         Caption         =   "Associate the &WAV and MID file formats with this Player?"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5295
      End
      Begin VB.CheckBox WMA_AVI_RMI 
         Caption         =   "Associate the W&MA, AVI and RMI file formats with this Player?"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   660
         Width           =   5295
      End
      Begin VB.CheckBox MPEG 
         Caption         =   "Associate MP&EG file formats with this player?"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1020
         Width           =   5295
      End
      Begin VB.CheckBox MXT 
         Caption         =   "Associate MX&T and VMB playlists with this player?"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1740
         Value           =   1  'Checked
         Width           =   5295
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Register Editor"
      Height          =   375
      Left            =   4200
      TabIndex        =   1
      Top             =   2460
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Associate"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2460
      Width           =   1395
   End
End
Attribute VB_Name = "frmFA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Dim Z As Boolean

Z = DeAssociateFile("wav", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("mid", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("wma", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("avi", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("rmi", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("mp3", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("mp2", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("mp1", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("mxt", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
Z = DeAssociateFile("vmb", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")

If Wav_Mid.Value = 1 Then Z = AssociateFile("wav", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If Wav_Mid.Value = 1 Then Z = AssociateFile("mid", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If WMA_AVI_RMI.Value = 1 Then Z = AssociateFile("wma", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If WMA_AVI_RMI.Value = 1 Then Z = AssociateFile("avi", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If WMA_AVI_RMI.Value = 1 Then Z = AssociateFile("rmi", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If MPEG.Value = 1 Then Z = AssociateFile("mp3", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If MPEG.Value = 1 Then Z = AssociateFile("mp2", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If MPEG.Value = 1 Then Z = AssociateFile("mp1", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If MXT.Value = 1 Then Z = AssociateFile("mxt", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")
If MXT.Value = 1 Then Z = AssociateFile("vmb", Chr$(34) + LowPath(App.path) + "vmbxt401.exe" + Chr$(34), "װאיכû Virtual MegaBox")

End
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim Z As Double
Z = Shell("regedit.exe", vbNormalFocus)
End Sub





Private Sub Form_Load()
On Error Resume Next
Open LowPath(App.path) + "vmbxt401.exe" For Input As #1: Close #1
If Err Then
  MsgBox "Error! The Quadrosoft Vitrual MegaBox XT 4.01 not installed or installed not correctly", vbCritical, "QS VMBXT 4.01 Setup"
  End
End If

End Sub


