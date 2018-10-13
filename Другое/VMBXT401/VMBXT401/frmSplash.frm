VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4245
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   7635
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   13  'Arrow and Hourglass
   ScaleHeight     =   4245
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Height          =   4050
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   7380
      Begin VB.Timer Timer1 
         Interval        =   500
         Left            =   4500
         Top             =   3240
      End
      Begin vmbxt.MTrack MTrack1 
         Height          =   375
         Left            =   5460
         TabIndex        =   8
         Top             =   2040
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   661
      End
      Begin VB.Label lbEDIT 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "РЕДАКЦІЯ"
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
         Left            =   6480
         TabIndex        =   12
         Top             =   2220
         Width           =   720
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   12
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   270
         Left            =   6195
         TabIndex        =   3
         Top             =   2760
         Width           =   1050
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Professional Audio Organizer For Windows"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   2760
         Width           =   4215
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Під дахом Васильківського Заводу Газового Обладнання, смт. Калинівка+Home Edition 01"
         ForeColor       =   &H0000FFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3540
         Width           =   3555
      End
      Begin VB.Image Image1 
         Height          =   1005
         Left            =   3540
         Picture         =   "frmSplash.frx":000C
         Stretch         =   -1  'True
         Top             =   240
         Width           =   3645
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Автор: ROMAN GEMINI (20061985)"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   2100
         TabIndex        =   7
         Top             =   2040
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Downloading playlist entries..."
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   4020
         TabIndex        =   6
         Top             =   3780
         Width           =   3240
      End
      Begin VB.Image imgLogo 
         Height          =   2325
         Left            =   120
         Picture         =   "frmSplash.frx":E5DA
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1995
      End
      Begin VB.Label lblCopyright 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright 2000-2004 (C) BRK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   4620
         TabIndex        =   2
         Top             =   3060
         Width           =   2595
      End
      Begin VB.Label lblCompany 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "RENNSoft ™"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Left            =   4320
         TabIndex        =   1
         Top             =   3270
         Width           =   2895
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Spring Edition"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   15.75
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   345
         Left            =   4470
         TabIndex        =   4
         Top             =   2460
         Width           =   2730
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual MegaBox"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   32.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FF80&
         Height          =   780
         Index           =   0
         Left            =   2200
         TabIndex        =   5
         Top             =   1240
         Width           =   4335
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "Virtual MegaBox"
         BeginProperty Font 
            Name            =   "Impact"
            Size            =   32.25
            Charset         =   204
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008080&
         Height          =   780
         Index           =   1
         Left            =   2230
         TabIndex        =   11
         Top             =   1280
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Form_Load()

On Error Resume Next
frmWCC.AddLogRecord ("Завантаження привітального вікна...")
Dim Ass As Integer, Exl As Long


frmWCC.AddLogRecord ("Перевірка інстанції...")
If App.PrevInstance = True Then MsgBox "Program is allready running!", vbCritical: End

frmWCC.AddLogRecord ("Самообстеження...")
Open LowPath(App.Path) + App.EXEName For Input As #1: Exl = LOF(1): Close #1

frmWCC.AddLogRecord ("Показ привітального вікна...")
Me.Show

frmWCC.AddLogRecord ("Облицьовування...")
Label5.Caption = "Professional Audio Organizer" + vbCrLf + "Header Size:     1.2M" + vbCrLf + "Programm Size: " + Format(Exl / 1000, "### ##0.0##") + "K"

lblVersion.Caption = "Версія XT " + GetVersion

frmWCC.AddLogRecord ("Перевіряю дисплей...")
If Screen.Height / Screen.TwipsPerPixelY < 576 Then
 MsgBox "Соррі, 720x576 дисплей мінімум!", vbExclamation, "Sorry! " + Format(Screen.Width / Screen.TwipsPerPixelX) + "x" + Format(Screen.Height / Screen.TwipsPerPixelY)
 End
Else
 For Ass = 0 To App.Revision Step App.Revision / 16
  MTrack1.Track = Ass
 Next
 MTrack1.Track = App.Revision
 Timer1.Enabled = True
End If

frmWCC.AddLogRecord ("Завантаження привітального вікна закінчено...")

End Sub

Private Sub imgLogo_DblClick()
Timer1.Enabled = Not Timer1.Enabled
If Timer1.Enabled = False Then Me.Label1.Caption = "На паузі"
If Timer1.Enabled = True Then Me.Label1.Caption = "Продовжую..."
End Sub


Private Sub Timer1_Timer()
frmWCC.AddLogRecord ("Завантаження програвача...")
wndMain.Show
frmWCC.AddLogRecord ("Вивантаження привітального вікна...")
Unload Me
End Sub


