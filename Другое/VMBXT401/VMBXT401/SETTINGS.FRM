VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Settings 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programm Settings"
   ClientHeight    =   6705
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   9885
   Icon            =   "settings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmTitle 
      Caption         =   "Title Formatting"
      Height          =   3735
      Left            =   4380
      TabIndex        =   34
      Top             =   420
      Width           =   2355
      Begin VB.TextBox txttagless 
         Height          =   285
         Left            =   180
         TabIndex        =   41
         ToolTipText     =   "Tagless title format"
         Top             =   600
         Width           =   1995
      End
      Begin VB.CommandButton btnReset 
         Caption         =   "Reset"
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   960
         Width           =   975
      End
      Begin VB.CommandButton btnDefault 
         Caption         =   "Back"
         Height          =   255
         Left            =   180
         TabIndex        =   39
         Top             =   960
         Width           =   1035
      End
      Begin VB.TextBox txtTtlFrm 
         Height          =   285
         Left            =   180
         TabIndex        =   36
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Special symbol list"
         Height          =   2175
         Left            =   60
         TabIndex        =   38
         Top             =   1500
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Spec. symbols:"
         Height          =   255
         Left            =   60
         TabIndex        =   37
         Top             =   1260
         Width           =   1215
      End
   End
   Begin VB.Frame frmStop 
      Caption         =   "Play completed mode"
      Height          =   2235
      Left            =   1980
      TabIndex        =   30
      Top             =   4200
      Width           =   2355
      Begin VB.OptionButton opOff 
         Caption         =   "Exit && Turn Off Computer"
         Height          =   195
         Left            =   120
         TabIndex        =   33
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton opExPlay 
         Caption         =   "Exit Player"
         Height          =   195
         Left            =   120
         TabIndex        =   32
         Top             =   780
         Width           =   2115
      End
      Begin VB.OptionButton opStOnly 
         Caption         =   "Stop Only"
         Height          =   195
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Value           =   -1  'True
         Width           =   2115
      End
   End
   Begin VB.PictureBox bButs 
      BackColor       =   &H00400000&
      Height          =   975
      Left            =   0
      ScaleHeight     =   915
      ScaleWidth      =   1815
      TabIndex        =   23
      Top             =   3300
      Width           =   1875
      Begin VB.CommandButton OKButton 
         Caption         =   "OK"
         Height          =   375
         Left            =   60
         TabIndex        =   25
         Top             =   60
         Width           =   1695
      End
      Begin VB.CommandButton CancelButton 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   60
         TabIndex        =   24
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2835
      ScaleWidth      =   1815
      TabIndex        =   22
      Top             =   420
      Width           =   1875
      Begin VB.TextBox txtPause 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00400000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   300
         TabIndex        =   43
         Text            =   "1000"
         Top             =   2580
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ms"
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   44
         Top             =   2580
         Width           =   195
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Pause between files:"
         ForeColor       =   &H0000FF00&
         Height          =   195
         Left            =   0
         TabIndex        =   42
         Top             =   2340
         Width           =   1470
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Formatting"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   35
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Stop Mode"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   29
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Playlist settings"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   28
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Title Window Settings"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   27
         Top             =   540
         Width           =   1545
      End
      Begin VB.Label lblIndex 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Standart settings"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   204
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   26
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.Frame frmPList 
      Caption         =   "Playlist Saving Settings"
      Height          =   2295
      Left            =   6780
      TabIndex        =   15
      Top             =   1860
      Width           =   2355
      Begin VB.CheckBox chCP 
         Caption         =   "Cursor Position"
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1800
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chTS 
         Caption         =   "Timer Settings"
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chPM 
         Caption         =   "Play Mode"
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1080
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chPP 
         Caption         =   "Play Position"
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   720
         Value           =   1  'Checked
         Width           =   1995
      End
      Begin VB.CheckBox chFS 
         Caption         =   "File Selection"
         Height          =   315
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Value           =   1  'Checked
         Width           =   1995
      End
   End
   Begin MSComDlg.CommonDialog cdTC 
      Left            =   6840
      Top             =   4620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Color           =   49152
      DialogTitle     =   "Select The Title Color"
   End
   Begin VB.Frame frmWind 
      Caption         =   "Title Windows Settings"
      Height          =   1335
      Left            =   6780
      TabIndex        =   9
      Top             =   420
      Width           =   2355
      Begin VB.CheckBox cmdFluent 
         Caption         =   "Fluent Scrolling"
         Height          =   255
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   600
         Width           =   1995
      End
      Begin VB.PictureBox ttlColor 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         Height          =   315
         Left            =   1680
         ScaleHeight     =   255
         ScaleWidth      =   435
         TabIndex        =   12
         Top             =   900
         Width           =   495
         Begin VB.Label lbTtlClr 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ABC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   204
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   30
            TabIndex        =   14
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "<<..Title Color..>>"
         Height          =   315
         Left            =   180
         TabIndex        =   11
         Top             =   900
         Width           =   1515
      End
      Begin VB.CheckBox cmdScroll 
         Caption         =   "Enable Scrolling"
         Height          =   255
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   300
         Width           =   1995
      End
   End
   Begin VB.Frame frmPlay 
      Caption         =   "Standart settings"
      Height          =   3735
      Left            =   1980
      TabIndex        =   0
      Top             =   420
      Width           =   2355
      Begin VB.CheckBox fndlost 
         Caption         =   "Seeker for lost files"
         Height          =   195
         Left            =   180
         TabIndex        =   45
         Top             =   3480
         Width           =   2055
      End
      Begin VB.Frame Frame2 
         Caption         =   "Repeat Mode"
         Height          =   1575
         Left            =   120
         TabIndex        =   4
         Top             =   1860
         Width           =   2115
         Begin VB.OptionButton opFileRep 
            Caption         =   "File Repeat"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1020
            Width           =   1755
         End
         Begin VB.OptionButton opListRep 
            Caption         =   "List Repeat"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   660
            Width           =   1755
         End
         Begin VB.OptionButton opNorep 
            Caption         =   "No Repeat"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   300
            Width           =   1755
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Play mode"
         Height          =   1575
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   2115
         Begin VB.OptionButton opStop 
            Caption         =   "Stop"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   1020
            Width           =   1755
         End
         Begin VB.OptionButton opDown 
            Caption         =   "Backward playing"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   660
            Width           =   1755
         End
         Begin VB.OptionButton opUp 
            Caption         =   "Forward playing"
            Height          =   375
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   300
            Width           =   1755
         End
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Player Settings Window"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Global Settings
Dim LclPlayMode As Integer
Dim LclRepeat As Integer

' Title Windows Settings
Dim LclWndScroll As Integer
Dim LclWndFluent As Integer
Dim LclWndColor As OLE_COLOR

' Playlist Saving Settings
Dim LclCHFS As Integer
Dim LclCHPP As Integer
Dim LclCHPM As Integer
Dim LclCHTS As Integer
Dim LclCHCP As Integer

' Stop Mode
Dim STM As Integer

'Title
Dim LclTitle As String
Dim LclTagLess As String
Dim lclSeeker As Integer

Option Explicit

Private Sub btnDefault_Click()
txtTtlFrm.Text = GlbTitle
txttagless.Text = GlbTagLess
End Sub


Private Sub btnReset_Click()
txtTtlFrm.Text = "%5 - %4 (%1.%3)"
txttagless.Text = "%4 (%1.%3)"
End Sub


Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub chCP_Click()
LclCHCP = chCP.Value

End Sub

Private Sub chFS_Click()
LclCHFS = chFS.Value
End Sub

Private Sub chPM_Click()
LclCHPM = chPM.Value

End Sub

Private Sub chPP_Click()
LclCHPP = chPP.Value

End Sub

Private Sub chTS_Click()
LclCHTS = chTS.Value

End Sub

Private Sub cmdColor_Click()
   On Error Resume Next
   
   cdTC.Color = lbTtlClr.ForeColor
   cdTC.ShowColor
   
   If Err Then Exit Sub
   
   lbTtlClr.ForeColor = cdTC.Color
   LclWndColor = cdTC.Color

End Sub

Private Sub cmdFluent_Click()
LclWndFluent = cmdFluent.Value
End Sub

Private Sub cmdScroll_Click()
LclWndScroll = cmdScroll.Value
End Sub

Private Sub fndlost_Click()
lclSeeker = fndlost.Value
End Sub

Private Sub Form_Load()
Me.Width = 4530
Label3.Caption = ""
Label3.Caption = Label3.Caption + "%1 - Filename" + vbCrLf
Label3.Caption = Label3.Caption + "%2 - Filepath" + vbCrLf
Label3.Caption = Label3.Caption + "%3 - File ext. (w/o point)" + vbCrLf
Label3.Caption = Label3.Caption + "%TAB - Insert a TAB symbol" + vbCrLf
Label3.Caption = Label3.Caption + "%4 - Title" + vbCrLf
Label3.Caption = Label3.Caption + "%5 - Artist" + vbCrLf
Label3.Caption = Label3.Caption + "%6 - Album" + vbCrLf
Label3.Caption = Label3.Caption + "%7 - Year" + vbCrLf
Label3.Caption = Label3.Caption + "%8 - Gerne" + vbCrLf
Label3.Caption = Label3.Caption + "%9 - File type" + vbCrLf
Label3.Caption = Label3.Caption + "%10 - File Size"

lblIndex_Click 0

'--------------------
Me.Show
On Error Resume Next
Dim Y, MY

MY = (ttlColor.Width + ttlColor.Height) / Screen.TwipsPerPixelY
For Y = 0 To MY
 ttlColor.Line (-10, Y * Screen.TwipsPerPixelY)-(Y * Screen.TwipsPerPixelX, -10), RGB(100 + (100 / MY * Y), 70 + (100 / MY * Y), 0)
Next

Select Case GlbPlayMode
 Case 0: opDown.Value = True
 Case 1: opUp.Value = True
 Case 2: opStop.Value = True
End Select
LclPlayMode = GlbPlayMode

Select Case GlbRepeat
 Case 0: opNorep.Value = True
 Case 1: opListRep.Value = True
 Case 2: opFileRep.Value = True
End Select
LclRepeat = GlbRepeat

Select Case GlbSTM
 Case 0: Me.opStOnly.Value = True
 Case 1: opExPlay.Value = True
 Case 2: opOff.Value = True
End Select
STM = GlbSTM

cmdScroll.Value = GlbWndScroll
cmdFluent.Value = GlbWndFluent
lbTtlClr.ForeColor = GlbWndColor

LclWndColor = GlbWndColor
LclWndScroll = GlbWndScroll
LclWndScroll = GlbWndColor

chFS.Value = GlbCHFS
chPP.Value = GlbCHPP
chPM.Value = GlbCHPM
chTS.Value = GlbCHTS
chCP.Value = GlbCHCP

LclCHFS = GlbCHFS
LclCHPP = GlbCHPP
LclCHPM = GlbCHPM
LclCHTS = GlbCHTS
LclCHCP = GlbCHCP
LclTitle = GlbTitle
LclTagLess = GlbTagLess
lclSeeker = GlbSeeker

txtTtlFrm.Text = LclTitle
txttagless.Text = LclTagLess

txtPause.Text = Format(GlbPause, "0")

fndlost.Value = lclSeeker

End Sub

Private Sub lblIndex_Click(Index As Integer)

  If Index = 4 Then frmTitle.Visible = True Else frmTitle.Visible = False
  If Index = 3 Then frmStop.Visible = True Else frmStop.Visible = False
  If Index = 2 Then frmPList.Visible = True Else frmPList.Visible = False
  If Index = 1 Then frmWind.Visible = True Else frmWind.Visible = False
  If Index = 0 Then frmPlay.Visible = True Else frmPlay.Visible = False
  
  If Index = 0 Then frmPlay.Top = 420: frmPlay.Left = 1980
  If Index = 1 Then frmWind.Top = 420: frmWind.Left = 1980
  If Index = 2 Then frmPList.Top = 420: frmPList.Left = 1980
  If Index = 3 Then frmStop.Top = 420: frmStop.Left = 1980
  If Index = 4 Then frmTitle.Top = 420: frmTitle.Left = 1980

  Dim X

  For X = 0 To lblIndex.Count - 1
    If X = Index Then lblIndex(X).FontBold = True Else lblIndex(X).FontBold = False
  Next

End Sub

Private Sub OKButton_Click()

GlbPlayMode = LclPlayMode
GlbRepeat = LclRepeat

' Title Windows Settings
GlbWndScroll = LclWndScroll
GlbWndFluent = LclWndFluent
GlbWndColor = LclWndColor

' Playlist Saving Settings
GlbCHFS = LclCHFS
GlbCHPP = LclCHPP
GlbCHPM = LclCHPM
GlbCHTS = LclCHTS
GlbCHCP = LclCHCP

GlbSTM = STM
GlbTitle = LclTitle
GlbTagLess = LclTagLess
GlbPause = Val(txtPause.Text)
GlbSeeker = lclSeeker


Setting.GlbSaveSettings
wndMain.ApplySettings
wndTitle.wndLS

Unload Me

End Sub

Private Sub opDown_Click()
LclPlayMode = 0
End Sub

Private Sub opExPlay_Click()
STM = 1
End Sub

Private Sub opFileRep_Click()
LclRepeat = 2
End Sub

Private Sub opListRep_Click()
LclRepeat = 1
End Sub

Private Sub opNorep_Click()
LclRepeat = 0
End Sub

Private Sub opOff_Click()
STM = 2
End Sub

Private Sub opStOnly_Click()
STM = 0
End Sub

Private Sub opStop_Click()
LclPlayMode = 2
End Sub

Private Sub opUp_Click()
LclPlayMode = 1
End Sub

Private Sub txtPause_Change()
If Val(txtPause.Text) > 5000 Then txtPause.Text = "5000"
End Sub

Private Sub txttagless_Change()
LclTagLess = txttagless.Text
End Sub

Private Sub txtTtlFrm_Change()
LclTitle = txtTtlFrm.Text
End Sub


