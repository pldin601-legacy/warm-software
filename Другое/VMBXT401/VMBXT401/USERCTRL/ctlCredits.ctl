VERSION 5.00
Begin VB.UserControl Credits 
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6825
   ScaleHeight     =   2340
   ScaleWidth      =   6825
   Begin VB.PictureBox picDest 
      BackColor       =   &H00000000&
      Height          =   1995
      Left            =   120
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   512
      TabIndex        =   1
      Top             =   180
      Width           =   3255
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1995
      Left            =   3420
      ScaleHeight     =   256
      ScaleMode       =   0  'User
      ScaleWidth      =   512
      TabIndex        =   0
      Top             =   180
      Visible         =   0   'False
      Width           =   3255
   End
End
Attribute VB_Name = "Credits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub DrawPicture(Text As String)
 
 
 picSource.Cls
 picSource.BackColor = RGB(255 * Rnd, 255 * Rnd, 255 * Rnd)
 picSource.CurrentX = 0
 picSource.CurrentY = 60
 picSource.Print Text
 
 Dim X
 
 For X = 0 To 100000
  n = Rnd * picDest.Height
  BitBlt picDest.hDC, 0, n, picDest.Width, 1, picSource.hDC, 0, n, vbSrcCopy
  Sleep 0.0001
 Next
 
End Sub

