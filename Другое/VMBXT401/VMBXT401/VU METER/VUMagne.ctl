VERSION 5.00
Begin VB.UserControl VUMG 
   BackColor       =   &H00000000&
   ClientHeight    =   2430
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   Picture         =   "VUMagne.ctx":0000
   ScaleHeight     =   162
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   241
   Begin VB.PictureBox picPeak 
      Appearance      =   0  'Flat
      BackColor       =   &H00000080&
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   1260
      ScaleHeight     =   105
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   720
      Width           =   195
   End
   Begin VB.Timer tmrFade 
      Left            =   2580
      Top             =   120
   End
   Begin VB.Image Image1 
      Height          =   150
      Left            =   540
      Picture         =   "VUMagne.ctx":088E
      Top             =   780
      Width           =   450
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0080C0FF&
      BorderWidth     =   2
      X1              =   52
      X2              =   12
      Y1              =   60
      Y2              =   28
   End
End
Attribute VB_Name = "VUMG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim LastSignL As Integer
Dim LastSignR As Integer

Dim CurrSignL As Integer
Dim CurrSignR As Integer


Sub SendSingnal(LeftChn As Integer)
CurrSignL = LeftChn
End Sub

Sub SLR(LeftChn As Integer)

Dim SMS As Currency
SMS = (2 / 100 * LeftChn) - 1

Line1.X2 = Line1.X1 + -Sin(-SMS) * 50
Line1.Y2 = Line1.Y1 + -Cos(-SMS) * 50

If LeftChn > 80 Then
 picPeak.BackColor = RGB(255, 0, 0)
Else
 picPeak.BackColor = RGB(100, 0, 0)
End If

End Sub


Sub VU_ON()
tmrFade.Interval = 10
End Sub

Private Sub Timer1_Timer()
Line1.X2 = Line1.X1 + -Sin(1) * 700
Line1.Y2 = Line1.Y1 + -Cos(1) * 700

End Sub

Private Sub tmrFade_Timer()

If CurrSignL = LastSignL Then
 LastSignL = CurrSignL
Else
 If CurrSignL > LastSignL Then
  LastSignL = LastSignL + 2
 End If
 If CurrSignL < LastSignL Then
  LastSignL = LastSignL - 1
 End If
End If

SLR LastSignL

End Sub


