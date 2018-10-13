VERSION 5.00
Begin VB.Form frmFind 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Знайти файл"
   ClientHeight    =   5790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5460
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrSearch 
      Interval        =   1000
      Left            =   1500
      Top             =   5280
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Пошук"
      Height          =   435
      Left            =   60
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin VB.CommandButton cmdJump 
      Caption         =   "Стрибати!"
      Height          =   435
      Left            =   4080
      TabIndex        =   5
      Top             =   5280
      Width           =   1275
   End
   Begin VB.CommandButton cmdCanc 
      Cancel          =   -1  'True
      Caption         =   "Відмінити"
      Height          =   435
      Left            =   2700
      TabIndex        =   4
      Top             =   5280
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Зміст:"
      Height          =   4215
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   5295
      Begin VB.ListBox lstEntr 
         Height          =   3765
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   5055
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Шукати по шаблону:"
      Height          =   855
      Left            =   60
      TabIndex        =   0
      Top             =   4320
      Width           =   5295
      Begin VB.TextBox txtSearch 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LFI As Integer
Private Sub cmdCanc_Click()
Unload Me
End Sub


Private Sub cmdJump_Click()
wndMain.lstSec.ListIndex = Me.lstEntr.ListIndex
Unload Me
End Sub

Private Sub cmdNext_Click()
 
 Dim Z As Integer, B As Integer, X As Integer
 X = lstEntr.ListIndex + 1
 
 For Z = X To Me.lstEntr.ListCount - 1
  B = InStr(1, LCase(lstEntr.List(Z)), LCase(Me.txtSearch.Text))
  If B > 0 Then lstEntr.ListIndex = Z: Exit For
 Next Z

 If B = 0 Then cmdNext.Enabled = False

End Sub

Private Sub Form_Load()
Dim A

For A = 0 To wndMain.lstSec.ListCount - 1
  Me.lstEntr.List(A) = wndMain.lstSec.List(A)
Next A

Me.lstEntr.ListIndex = 0

End Sub


Private Sub tmrSearch_Timer()
If LFI = -1 Then Exit Sub

LFI = LFI + 1

If LFI = 2 Then
 If txtSearch.Text <> "" Then
 Dim Z As Integer, B As Integer
 For Z = 0 To Me.lstEntr.ListCount - 1
  B = InStr(1, LCase(lstEntr.List(Z)), LCase(Me.txtSearch.Text))
  If B > 0 Then lstEntr.ListIndex = Z: Exit For
 Next Z
 End If
LFI = -1
End If

End Sub

Private Sub txtSearch_Change()
LFI = 0
If txtSearch.Text = "" Then cmdNext.Enabled = False
If txtSearch.Text <> "" Then cmdNext.Enabled = True
End Sub


