VERSION 5.00
Begin VB.Form wpPlayer 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   Caption         =   "WinPly MP3 Player 1.0"
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "wpPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox seeker 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   180
      ScaleHeight     =   1
      ScaleMode       =   0  'User
      ScaleWidth      =   150
      TabIndex        =   11
      Top             =   1620
      Width           =   5235
   End
   Begin VB.CheckBox btnPL 
      Caption         =   "PLAYLIST"
      Height          =   255
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin vmbxt.BigTime TimeX 
      CausesValidation=   0   'False
      Height          =   360
      Left            =   900
      TabIndex        =   9
      Top             =   480
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   635
   End
   Begin vmbxt.QSImgButton btnExit 
      Height          =   255
      Left            =   5160
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Picture         =   "wpPlayer.frx":030A
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnMinimize 
      Height          =   255
      Left            =   4920
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Picture         =   "wpPlayer.frx":07FA
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnEject 
      Height          =   195
      Left            =   3180
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   344
      Picture         =   "wpPlayer.frx":0C9E
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnNext 
      Height          =   195
      Left            =   2400
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   344
      Picture         =   "wpPlayer.frx":116A
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnStop 
      Height          =   195
      Left            =   1920
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   344
      Picture         =   "wpPlayer.frx":1682
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnPause 
      Height          =   195
      Left            =   1440
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   344
      Picture         =   "wpPlayer.frx":1B5C
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnPlay 
      Height          =   195
      Left            =   960
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   344
      Picture         =   "wpPlayer.frx":207E
      BackColor       =   0
   End
   Begin vmbxt.QSImgButton btnPrev 
      Height          =   195
      Left            =   480
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   344
      Picture         =   "wpPlayer.frx":2548
      BackColor       =   0
   End
   Begin VB.Image RTMODE 
      Height          =   405
      Left            =   240
      Picture         =   "wpPlayer.frx":2A60
      Top             =   420
      Width           =   465
   End
   Begin VB.Label btSettings 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   900
      Width           =   135
   End
   Begin VB.Label st_mn 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "FULL STEREO"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QèA!)ROS()FÜ Virtual Remote Control 1.4"
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
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Tag             =   "DOWN"
      Top             =   60
      Width           =   5235
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "MCI MUSIC WORLD"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   180
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TRACK"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4740
      TabIndex        =   4
      Top             =   660
      Width           =   600
   End
   Begin VB.Label TRK 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3900
      TabIndex        =   3
      Top             =   660
      Width           =   735
   End
   Begin VB.Label TOTL 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2340
      TabIndex        =   2
      Top             =   660
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "LENGTH"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Index           =   0
      Left            =   3060
      TabIndex        =   1
      Top             =   660
      Width           =   720
   End
   Begin VB.Label lblTrackTitle 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2220
      TabIndex        =   0
      Top             =   360
      Width           =   3135
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      Height          =   795
      Left            =   180
      Top             =   360
      Width           =   1935
   End
   Begin VB.Shape Shape2 
      Height          =   1275
      Left            =   120
      Top             =   240
      Width           =   5355
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00E0E0E0&
      X1              =   60
      X2              =   5520
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   60
      X2              =   5520
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00E0E0E0&
      Height          =   2175
      Left            =   60
      Top             =   180
      Width           =   5475
   End
End
Attribute VB_Name = "wpPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX, OldY As Integer


Private Sub btnEject_Click()

wndAddFiles.lstFile.Tag = "Open"
wndAddFiles.Show 1, Me


wndAddFiles.lstFile.Tag = ""

End Sub

Private Sub btnExit_Click()
Me.Hide
wndMain.Show
End Sub

Private Sub btnMinimize_Click()
Me.WindowState = 1
End Sub

Private Sub btnNext_Click()
wndMain.Command_Next
End Sub

Private Sub btnPause_Click()
wndMain.Command_Pause
End Sub

Private Sub btnPL_Click()
wpPlaylist.Visible = btnPL.Value

End Sub

Private Sub btnPlay_Click()

If wndMain.MMHeader.Mode = 529 Then
 wndMain.Command_Pause
Else
 wndMain.Command_Play
End If

End Sub

Private Sub btnPrev_Click()
wndMain.Command_Prev
End Sub

Private Sub btnStop_Click()
wndMain.ResetStatus
End Sub

Private Sub Form_Load()
Me.Scale (0, 0)-(640, 264)

Dim a As Integer, CL As Integer
  
For a = 0 To 264
  CL = 85 + ((a / 6) * Sin(a))
  Me.Line (0, a)-(640, a), RGB(CL, CL, CL), BF
Next

Line (0, 0)-(Me.Width, 0), RGB(200, 200, 200)
Line (0, 0)-(0, Me.Height), RGB(200, 200, 200)
Line (Me.Width - 15, 0)-(Me.Width - 15, Me.Height - 15), RGB(100, 100, 100)
Line (0, Me.Height - 15)-(Me.Width - 15, Me.Height - 15), RGB(100, 100, 100)

Load wpPlaylist

End Sub

Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X
OldY = Y
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ML, MT

If Button = 1 Then

 ML = Me.Left + (X - OldX)
 MT = Me.Top + (Y - OldY)
 
 If ML <= 120 Then ML = 0
 If MT <= 120 Then MT = 0
 
 If ML > Screen.Width - Me.Width - 120 Then ML = Screen.Width - Me.Width
 If MT > Screen.Height - Me.Height - 120 Then MT = Screen.Height - Me.Height
 
 Me.Left = ML
 Me.Top = MT

 X = ML
 Y = MT
 
End If

End Sub

Private Sub seeker_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 wndMain.Scr_Down
End Sub

Private Sub seeker_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
 
 If Button = 1 Then
   wndMain.Scr_Scroll (X)
   SeekChange seeker.ScaleWidth, CInt(X)
 End If
 
End Sub


Sub SeekChange(Max As Integer, Min As Integer)
Dim MOCbKA As Integer
Dim Valve As Integer
Dim Vise As Integer

Vise = 150 / Max * Min
seeker.Cls

For MOCbKA = Vise To Vise + 10
 Valve = MOCbKA - Vise
 seeker.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 - (255 / 10 * Valve), 255 - (255 / 10 * Valve), 0), BF
Next

For MOCbKA = Vise - 10 To Vise
 Valve = MOCbKA - (Vise - 10)
 seeker.Line (MOCbKA, 0)-(MOCbKA + 1, 1), RGB(255 / 10 * Valve, 255 / 10 * Valve, 0), BF
Next

seeker.Line (Vise, 0)-(Vise + 1, 1), RGB(0, 255, 0), BF

End Sub

Private Sub seeker_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

 wndMain.Scr_Up (X)

End Sub


Private Sub seeker_Scroll()
 
End Sub


