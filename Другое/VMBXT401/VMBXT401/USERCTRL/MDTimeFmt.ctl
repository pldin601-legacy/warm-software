VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.UserControl MTimer 
   BackColor       =   &H00000000&
   CanGetFocus     =   0   'False
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ScaleHeight     =   1905
   ScaleWidth      =   2505
   Begin PicClip.PictureClip DigResource 
      Left            =   900
      Top             =   1140
      _ExtentX        =   1455
      _ExtentY        =   1164
      _Version        =   393216
      Rows            =   2
      Cols            =   5
      Picture         =   "MDTimeFmt.ctx":0000
   End
   Begin VB.Image Picture5 
      Height          =   300
      Left            =   360
      Picture         =   "MDTimeFmt.ctx":1D32
      Top             =   60
      Width           =   150
   End
   Begin VB.Image Dig1 
      Height          =   330
      Index           =   3
      Left            =   0
      Picture         =   "MDTimeFmt.ctx":1FF4
      Top             =   0
      Width           =   165
   End
   Begin VB.Image Dig1 
      Height          =   330
      Index           =   2
      Left            =   180
      Picture         =   "MDTimeFmt.ctx":234E
      Top             =   0
      Width           =   165
   End
   Begin VB.Image Dig1 
      Height          =   330
      Index           =   1
      Left            =   540
      Picture         =   "MDTimeFmt.ctx":26A8
      Top             =   0
      Width           =   165
   End
   Begin VB.Image Dig1 
      Height          =   330
      Index           =   0
      Left            =   720
      Picture         =   "MDTimeFmt.ctx":2A02
      Top             =   0
      Width           =   165
   End
   Begin VB.Image Image1 
      Height          =   330
      Left            =   1260
      Picture         =   "MDTimeFmt.ctx":2D5C
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
   Begin VB.Image pReset 
      Height          =   330
      Left            =   1080
      Picture         =   "MDTimeFmt.ctx":30B6
      Top             =   0
      Visible         =   0   'False
      Width           =   165
   End
End
Attribute VB_Name = "MTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Sub Off()
Dig1(0).Picture = Image1.Picture
Dig1(1).Picture = Image1.Picture
Dig1(2).Picture = Image1.Picture
Dig1(3).Picture = Image1.Picture
Picture5.Visible = False
End Sub

Public Property Let TimeSet(TimeNow As String)
On Error Resume Next
Dim Secs As String
Dim Mins As String

Secs = Mid$(TimeNow, 4, 2)
Mins = Mid$(TimeNow, 1, 2)

Picture5.Visible = True
Dig1(0).Picture = DigResource.GraphicCell(Val(Mid$(Secs, 2, 1)))
Dig1(1).Picture = DigResource.GraphicCell(Val(Mid$(Secs, 1, 1)))
Dig1(2).Picture = DigResource.GraphicCell(Val(Mid$(Mins, 2, 1)))
If Mid$(Mins, 1, 1) = "0" Then
 Dig1(3).Picture = Image1.Picture
Else
 Dig1(3).Picture = DigResource.GraphicCell(Val(Mid$(Mins, 1, 1)))
End If

End Property

Public Property Let TimeSet2(TimeNow As String)
Dim Secs As String
Dim Mins As String

Secs = Mid$(TimeNow, 4, 2)
Mins = Mid$(TimeNow, 1, 2)

Picture5.Visible = False
Dig1(0).Picture = DigResource.GraphicCell(Val(Mid$(Secs, 2, 1)))
Dig1(1).Picture = DigResource.GraphicCell(Val(Mid$(Secs, 1, 1)))
Dig1(2).Picture = DigResource.GraphicCell(Val(Mid$(Mins, 2, 1)))
If Mid$(Mins, 1, 1) = "0" Then
 Dig1(3).Picture = Image1.Picture
Else
 Dig1(3).Picture = DigResource.GraphicCell(Val(Mid$(Mins, 1, 1)))
End If

End Property

Public Sub Reset()
Dig1(0).Picture = pReset.Picture
Dig1(1).Picture = pReset.Picture
Dig1(2).Picture = pReset.Picture
Dig1(3).Picture = pReset.Picture
Picture5.Visible = False
End Sub

Private Sub UserControl_Resize()
UserControl.Height = Dig1(0).Height
UserControl.Width = Dig1(0).Width + Dig1(0).Width + Dig1(0).Width + Dig1(0).Width + Picture5.Width + 100
End Sub

Public Sub InitTimer()
For Digs = 9 To 0 Step -1
Dig1(0).Picture = DigResource.GraphicCell(Digs)
Dig1(1).Picture = DigResource.GraphicCell(Digs)
Dig1(2).Picture = DigResource.GraphicCell(Digs)
Dig1(3).Picture = DigResource.GraphicCell(Digs)
Sleep (0.25)
Next
Picture5.Visible = False
End Sub
Sub Sleep(tm As Currency)
tm2 = Timer
Do: DoEvents: Loop While Timer <> tm2 + tm
End Sub
