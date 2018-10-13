VERSION 5.00
Begin VB.Form rtc 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Virtual RTC"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2460
   Icon            =   "rtc.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   2460
   StartUpPosition =   3  'Windows Default
   Begin vmbxt401.QSImgButton btstop 
      Height          =   315
      Left            =   1980
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Picture         =   "rtc.frx":0442
      BackColor       =   0
   End
   Begin vmbxt401.QSImgButton btPause 
      Height          =   315
      Left            =   1500
      Top             =   960
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      Picture         =   "rtc.frx":0734
      BackColor       =   0
   End
   Begin vmbxt401.QSImgButton btplay 
      Height          =   315
      Left            =   1080
      Top             =   960
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   556
      Picture         =   "rtc.frx":097E
      BackColor       =   0
   End
   Begin vmbxt401.QSImgButton btNext 
      Height          =   315
      Left            =   1740
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Picture         =   "rtc.frx":0BD8
      BackColor       =   0
   End
   Begin vmbxt401.QSImgButton btPrev 
      Height          =   315
      Left            =   1080
      Top             =   1320
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      Picture         =   "rtc.frx":0F36
      BackColor       =   0
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   735
      Left            =   120
      ScaleHeight     =   675
      ScaleWidth      =   2175
      TabIndex        =   0
      Top             =   180
      Width           =   2235
      Begin vmbxt401.MTrack MTrack1 
         Height          =   315
         Left            =   60
         TabIndex        =   1
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   556
      End
      Begin vmbxt401.MTimer tmPos 
         Height          =   330
         Left            =   1200
         Top             =   240
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   582
      End
      Begin VB.Label idx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Position"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Index           =   1
         Left            =   1200
         TabIndex        =   3
         Top             =   60
         Width           =   540
      End
      Begin VB.Label idx 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Track"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   165
         Index           =   0
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.Image Image1 
      Height          =   405
      Left            =   120
      Picture         =   "rtc.frx":1308
      Top             =   960
      Width           =   465
   End
End
Attribute VB_Name = "rtc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
