VERSION 5.00
Begin VB.UserControl VU 
   BackColor       =   &H00000000&
   ClientHeight    =   2550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   Picture         =   "VU.ctx":0000
   ScaleHeight     =   2550
   ScaleWidth      =   3615
   Begin VB.Timer tmrFade 
      Left            =   2580
      Top             =   120
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   7
      Left            =   2040
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   15
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   6
      Left            =   1800
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   14
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   5
      Left            =   1560
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   13
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   4
      Left            =   1320
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   12
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   3
      Left            =   1080
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   11
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   2
      Left            =   840
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   10
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   1
      Left            =   600
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   9
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indLeft 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   0
      Left            =   360
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   8
      Top             =   480
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   7
      Left            =   2040
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   7
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00000080&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   6
      Left            =   1800
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   6
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00008080&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   5
      Left            =   1560
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   5
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   4
      Left            =   1320
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   4
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   3
      Left            =   1080
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   3
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   2
      Left            =   840
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   1
      Left            =   600
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   60
      Width           =   195
   End
   Begin VB.PictureBox indRight 
      BackColor       =   &H0000FF00&
      BorderStyle     =   0  'None
      Height          =   75
      Index           =   0
      Left            =   360
      ScaleHeight     =   75
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   60
      Width           =   195
   End
   Begin VB.Label ChnLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "L"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   17
      Top             =   420
      Width           =   135
   End
   Begin VB.Label ChnLabel 
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   16
      Top             =   0
      Width           =   135
   End
End
Attribute VB_Name = "VU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim LastSignL As Integer
Dim LastSignR As Integer

Dim CurrSignL As Integer
Dim CurrSignR As Integer


Sub SendSingnal(LeftChn As Integer, RightChn As Integer)
CurrSignL = LeftChn
CurrSignR = RightChn
tmrFade_Timer
End Sub

Private Sub SLR(LeftChn As Integer, RightChn As Integer)

If LeftChn > 0 Then indLeft(0).BackColor = RGB(5, 250, 5) Else indLeft(0).BackColor = RGB(5, 100, 5)
If LeftChn > 20 Then indLeft(1).BackColor = RGB(5, 250, 5) Else indLeft(1).BackColor = RGB(5, 100, 5)
If LeftChn > 30 Then indLeft(2).BackColor = RGB(5, 250, 5) Else indLeft(2).BackColor = RGB(5, 100, 5)
If LeftChn > 40 Then indLeft(3).BackColor = RGB(5, 250, 5) Else indLeft(3).BackColor = RGB(5, 100, 5)
If LeftChn > 50 Then indLeft(4).BackColor = RGB(5, 250, 5) Else indLeft(4).BackColor = RGB(5, 100, 5)
If LeftChn > 60 Then indLeft(5).BackColor = RGB(250, 250, 5) Else indLeft(5).BackColor = RGB(100, 100, 5)
If LeftChn > 80 Then indLeft(6).BackColor = RGB(250, 5, 5) Else indLeft(6).BackColor = RGB(100, 5, 5)
If LeftChn >= 99 Then indLeft(7).BackColor = RGB(250, 5, 5) Else indLeft(7).BackColor = RGB(100, 5, 5)

If RightChn > 0 Then indRight(0).BackColor = RGB(5, 250, 5) Else indRight(0).BackColor = RGB(5, 100, 5)
If RightChn > 20 Then indRight(1).BackColor = RGB(5, 250, 5) Else indRight(1).BackColor = RGB(5, 100, 5)
If RightChn > 30 Then indRight(2).BackColor = RGB(5, 250, 5) Else indRight(2).BackColor = RGB(5, 100, 5)
If RightChn > 40 Then indRight(3).BackColor = RGB(5, 250, 5) Else indRight(3).BackColor = RGB(5, 100, 5)
If RightChn > 50 Then indRight(4).BackColor = RGB(5, 250, 5) Else indRight(4).BackColor = RGB(5, 100, 5)
If RightChn > 60 Then indRight(5).BackColor = RGB(250, 250, 5) Else indRight(5).BackColor = RGB(100, 100, 5)
If RightChn > 80 Then indRight(6).BackColor = RGB(250, 5, 5) Else indRight(6).BackColor = RGB(100, 5, 5)
If RightChn >= 99 Then indRight(7).BackColor = RGB(250, 5, 5) Else indRight(7).BackColor = RGB(100, 5, 5)

End Sub


Sub VU_ON()
tmrFade.Interval = 10
End Sub

Private Sub tmrFade_Timer()

If CurrSignL >= LastSignL Then
 LastSignL = CurrSignL
Else
 LastSignL = LastSignL - 4
End If

If CurrSignR >= LastSignR Then
 LastSignR = CurrSignR
Else
 LastSignR = LastSignR - 4
End If

SLR LastSignL, LastSignR

End Sub


