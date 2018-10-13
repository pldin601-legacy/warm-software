VERSION 5.00
Begin VB.Form txtEditor 
   Caption         =   "Quadrosoft Virtual Text 1.0"
   ClientHeight    =   3795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   6075
   Icon            =   "frmWords.frx":0000
   LinkTopic       =   "Form1"
   NegotiateMenus  =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3795
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   3795
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   6075
   End
End
Attribute VB_Name = "txtEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Changed As Boolean
Dim LastFileName As String
Private Sub Form_Resize()
If Me.WindowState = 1 Then Exit Sub
Text1.Height = txtEditor.Height - 400
Text1.Width = txtEditor.Width - (8 * Screen.TwipsPerPixelY)
End Sub


Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Dim O
If Button = 2 Then
   O = FreeFile
   Open Me.Caption For Output As #O
   Print #O, Text1.Text
   Close #O
   If Err = 0 Then MsgBox ""
End If
End Sub


