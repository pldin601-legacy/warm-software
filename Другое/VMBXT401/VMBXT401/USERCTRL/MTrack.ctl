VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl MTrack 
   BackColor       =   &H00000000&
   BackStyle       =   0  'Transparent
   ClientHeight    =   2955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   2955
   ScaleWidth      =   3810
   Begin PicClip.PictureClip Error1 
      Left            =   1800
      Top             =   2160
      _ExtentX        =   873
      _ExtentY        =   582
      _Version        =   393216
      Cols            =   3
      Picture         =   "MTrack.ctx":0000
   End
   Begin PicClip.PictureClip LastR 
      Left            =   420
      Top             =   2040
      _ExtentX        =   1164
      _ExtentY        =   582
      _Version        =   393216
      Cols            =   4
      Picture         =   "MTrack.ctx":08EA
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   960
      Top             =   720
   End
   Begin PicClip.PictureClip PClip1 
      Left            =   840
      Top             =   1500
      _ExtentX        =   2037
      _ExtentY        =   582
      _Version        =   393216
      Cols            =   7
      Picture         =   "MTrack.ctx":1494
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   4
      Left            =   0
      Picture         =   "MTrack.ctx":28D6
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   4
      Top             =   0
      Width           =   165
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   3
      Left            =   180
      Picture         =   "MTrack.ctx":2C30
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   3
      Top             =   0
      Width           =   165
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   2
      Left            =   360
      Picture         =   "MTrack.ctx":2F8A
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   2
      Top             =   0
      Width           =   165
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   0
      Left            =   720
      Picture         =   "MTrack.ctx":32E4
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   1
      Top             =   0
      Width           =   165
   End
   Begin VB.PictureBox Dig1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000040&
      BorderStyle     =   0  'None
      Height          =   330
      Index           =   1
      Left            =   540
      Picture         =   "MTrack.ctx":363E
      ScaleHeight     =   330
      ScaleWidth      =   165
      TabIndex        =   0
      Top             =   0
      Width           =   165
   End
   Begin PicClip.PictureClip DigRes 
      Left            =   2460
      Top             =   1200
      _ExtentX        =   1455
      _ExtentY        =   1164
      _Version        =   393216
      Rows            =   2
      Cols            =   5
      Picture         =   "MTrack.ctx":3998
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   240
      Picture         =   "MTrack.ctx":56CA
      Top             =   840
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "MTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Sub LastRec()
Dig1(4).Picture = LastR.GraphicCell(0)
Dig1(3).Picture = LastR.GraphicCell(1)
Dig1(2).Picture = LastR.GraphicCell(2)
Dig1(1).Picture = LastR.GraphicCell(3)
Dig1(0).Picture = PClip1.GraphicCell(6)
End Sub

Public Property Let Track(Track As Integer)
Dim X As String, M, Z
Timer1.Enabled = False

M = Len(Str(Track))
X = Format$(Track, "00000")
Dig1(0).Picture = DigRes.GraphicCell(Val(Mid$(X, 5, 1)))
Dig1(1).Picture = DigRes.GraphicCell(Val(Mid$(X, 4, 1)))
Dig1(2).Picture = DigRes.GraphicCell(Val(Mid$(X, 3, 1)))
Dig1(3).Picture = DigRes.GraphicCell(Val(Mid$(X, 2, 1)))
Dig1(4).Picture = DigRes.GraphicCell(Val(Mid$(X, 1, 1)))

For Z = 5 To M Step -1
 Dig1(Z - 1).Picture = Image1.Picture
Next

End Property

Public Property Let TrackX(Track As Integer)
Dim X As String, M
Timer1.Enabled = False


X = Format$(Track, "00000")
Dig1(0).Picture = DigRes.GraphicCell(Val(Mid$(X, 5, 1)))
Dig1(1).Picture = DigRes.GraphicCell(Val(Mid$(X, 4, 1)))
Dig1(2).Picture = DigRes.GraphicCell(Val(Mid$(X, 3, 1)))
Dig1(3).Picture = DigRes.GraphicCell(Val(Mid$(X, 2, 1)))
If Val(Mid$(X, 1, 1)) = "0" Then Dig1(4).Picture = Image1.Picture Else Dig1(4).Picture = DigRes.GraphicCell(Val(Mid$(X, 1, 1)))

End Property

Sub UnWait()
Timer1.Enabled = False
End Sub

Sub Waiting()
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
Dim N
N = Fix((Timer * 10) / 5) Mod 14

Select Case N

Case 0
Dig1(4).Picture = PClip1.GraphicCell(0)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 1
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(0)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 2
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(0)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 3
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(0)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 4
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(0)

Case 5
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(2)

Case 6
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(3)

Case 7
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(1)

Case 8
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(1)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 9
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(1)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 10
Dig1(4).Picture = PClip1.GraphicCell(6)
Dig1(3).Picture = PClip1.GraphicCell(1)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 11
Dig1(4).Picture = PClip1.GraphicCell(1)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 12
Dig1(4).Picture = PClip1.GraphicCell(5)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

Case 13
Dig1(4).Picture = PClip1.GraphicCell(4)
Dig1(3).Picture = PClip1.GraphicCell(6)
Dig1(2).Picture = PClip1.GraphicCell(6)
Dig1(1).Picture = PClip1.GraphicCell(6)
Dig1(0).Picture = PClip1.GraphicCell(6)

End Select


End Sub
Sub ErrorFound()
Dig1(4).Picture = Error1.GraphicCell(0)
Dig1(3).Picture = Error1.GraphicCell(1)
Dig1(2).Picture = Error1.GraphicCell(1)
Dig1(1).Picture = Error1.GraphicCell(2)
Dig1(0).Picture = Error1.GraphicCell(1)
End Sub

