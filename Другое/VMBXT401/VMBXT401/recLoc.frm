VERSION 5.00
Begin VB.Form recLoc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Локалізація головних списків"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "recLoc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCancel 
      Caption         =   "Відмінити"
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   3840
      Width           =   1335
   End
   Begin VB.CommandButton btnOk 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.DriveListBox lstDrive 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   3420
      Width           =   4515
   End
   Begin VB.DirListBox lstDir 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   4515
   End
   Begin VB.Label lblPath 
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4515
   End
End
Attribute VB_Name = "recLoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCancel_Click()
Unload Me
End Sub

Private Sub btnOk_Click()
GlbRecentDir = lstDir.Path
Setting.GlbSaveSettings
Unload Me
End Sub


Private Sub Form_Load()
On Error Resume Next
lstDrive.Drive = Mid(GlbRecentDir, 1, 2)
lstDir.Path = GlbRecentDir
End Sub

Private Sub lstDir_Change()
lblPath.Caption = lstDir.Path
End Sub

Private Sub lstDrive_Change()

On Error Resume Next

lstDir.Path = lstDrive.Drive
If Err Then lstDrive.Drive = Mid(lstDir.Path, 1, 2)

End Sub


