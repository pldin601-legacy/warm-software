VERSION 5.00
Begin VB.Form cmdIns 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Play Completed Mode"
   ClientHeight    =   1950
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Enabled         =   0   'False
      Height          =   315
      Left            =   2580
      TabIndex        =   6
      Top             =   1320
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   180
      TabIndex        =   5
      Top             =   1320
      Width           =   2415
   End
   Begin VB.OptionButton opLaunch 
      Caption         =   "Launch Programm:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   1080
      Width           =   1755
   End
   Begin VB.OptionButton opCmdShut 
      Caption         =   "Shut Down System"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   660
      Width           =   2055
   End
   Begin VB.OptionButton opCmdExit 
      Caption         =   "Exit Player"
      Height          =   255
      Left            =   180
      TabIndex        =   2
      Top             =   240
      Width           =   1515
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "OK"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "cmdIns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()
On Error Resume Next

If opCmdExit.Value = True Then
  Me.Tag = "1"
End If

If opCmdShut.Value = True Then
  Me.Tag = "2"
End If

Unload Me

End Sub


