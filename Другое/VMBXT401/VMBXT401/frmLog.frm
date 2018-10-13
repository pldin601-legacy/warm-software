VERSION 5.00
Begin VB.Form frmWCC 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Вікно командного спілкування з плеєром"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   8715
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtUser 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   5220
      Width           =   6855
   End
   Begin VB.TextBox txtIndic 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   204
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5115
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "frmLog.frx":0000
      Top             =   60
      Width           =   8595
   End
End
Attribute VB_Name = "frmWCC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

' Initializing

txtIndic.Text = "RENNSOFT Multimedia Software (C) 2003" + vbCrLf
txtIndic.Text = txtIndic.Text + "Virtual MegaBox Версія " + GetVersion + vbCrLf
txtIndic.Text = txtIndic.Text + "Вікно віртуального спілкування з програмою " + vbCrLf
txtIndic.Text = txtIndic.Text + "---" + vbCrLf
txtIndic.Text = txtIndic.Text + "Завантаження процессора закінчено...ОК" + vbCrLf
txtIndic.Text = txtIndic.Text + "Старт програми" + vbCrLf
txtIndic.Text = txtIndic.Text + vbCrLf



End Sub

Private Sub Form_Unload(Cancel As Integer)
Cancel = 1
Me.Hide
End Sub


Sub AddLogRecord(dREc As String)

' txtIndic.Text = txtIndic.Text + dREc + vbCrLf

End Sub
