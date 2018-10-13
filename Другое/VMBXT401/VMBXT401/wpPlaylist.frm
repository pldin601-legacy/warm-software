VERSION 5.00
Begin VB.Form wpPlaylist 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   ClientHeight    =   2955
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   5475
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2955
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin vmbxt.QSImgButton btnCloseMe 
      Height          =   255
      Left            =   5040
      Top             =   60
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
      Enabled         =   0   'False
      Picture         =   "wpPlaylist.frx":0000
      BackColor       =   0
   End
   Begin VB.ListBox lstFiles 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   1860
      IntegralHeight  =   0   'False
      Left            =   180
      TabIndex        =   2
      Top             =   420
      Width           =   5115
   End
   Begin VB.PictureBox ctrls 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   5235
      TabIndex        =   1
      Top             =   2520
      Width           =   5235
      Begin vmbxt.QSImgButton btncopy 
         Height          =   195
         Left            =   4620
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Picture         =   "wpPlaylist.frx":04F0
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnKill 
         Height          =   195
         Left            =   4080
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Picture         =   "wpPlaylist.frx":09E6
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnRemove 
         Height          =   195
         Left            =   3360
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Picture         =   "wpPlaylist.frx":0EE8
         BackColor       =   0
      End
      Begin vmbxt.QSImgButton btnAdd 
         Height          =   195
         Left            =   2820
         Top             =   60
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   344
         Picture         =   "wpPlaylist.frx":1394
         BackColor       =   0
      End
      Begin vmbxt.QSButton btnMenu 
         Height          =   195
         Left            =   60
         Top             =   60
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   344
         BackColor       =   0
         Caption         =   "Playlist Menu"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Small Fonts"
         FontSize        =   6.75
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   1380
         TabIndex        =   3
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "QèA!)ROS()FÜ MegaBox Mini playlist"
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
      TabIndex        =   0
      Top             =   60
      Width           =   5115
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00E0E0E0&
      Height          =   2175
      Left            =   60
      Top             =   180
      Width           =   5355
   End
   Begin VB.Menu mnPLM 
      Caption         =   "Playlist Menu"
      Visible         =   0   'False
      Begin VB.Menu mnSavePl 
         Caption         =   "Save Playlist..."
      End
      Begin VB.Menu mnOpenPl 
         Caption         =   "Open Playlist..."
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnClear 
         Caption         =   "Clear Playlist"
      End
      Begin VB.Menu mnRemove 
         Caption         =   "Remove this item from playlist"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDelete 
         Caption         =   "Delete Files"
      End
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Files"
      End
   End
End
Attribute VB_Name = "wpPlaylist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX, OldY
Dim stdx, stdy

Private Sub btnCloseMe_Click()
Me.Visible = False
End Sub

Private Sub btnMenu_Click()
PopupMenu mnPLM
End Sub


Private Sub Form_Load()
stdx = Me.Width


End Sub


Private Sub Form_Resize()
Me.Width = stdx

If Me.Height < 2000 Then Me.Height = 2000

ctrls.Top = Me.Height - ctrls.Height - 200
Shape1.Height = Me.Height - ctrls.Height - 500
lstFiles.Height = Me.Height - ctrls.Height - 850


End Sub


Private Sub Label2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
OldX = X
OldY = Y
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then Me.Move Me.Left + (X - OldX), Me.Top + (Y - OldY)
If Button = 2 Then Me.Move 120 * Fix((Me.Left + (X - OldX)) / 120), 120 * Fix((Me.Top + (Y - OldY)) / 120)
End Sub


