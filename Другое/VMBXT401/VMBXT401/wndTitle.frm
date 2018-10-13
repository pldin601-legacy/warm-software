VERSION 5.00
Begin VB.Form wndTitle 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00808080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10710
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrScroll 
      Interval        =   500
      Left            =   6840
      Top             =   420
   End
   Begin VB.Timer tmrFluent 
      Interval        =   40
      Left            =   6300
      Top             =   420
   End
   Begin VB.Label lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RS"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   36
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   720
      Index           =   1
      Left            =   900
      MouseIcon       =   "wndTitle.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   915
   End
   Begin VB.Label lab 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "RS"
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   36
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   720
      Index           =   0
      Left            =   -60
      MouseIcon       =   "wndTitle.frx":0442
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   120
      UseMnemonic     =   0   'False
      Width           =   915
   End
End
Attribute VB_Name = "wndTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim OldX, OldY
Dim Sacred As Integer

Sub wndLS()

 If GlbWndScroll = 1 Then
  
  If GlbWndFluent = 1 Then
   tmrFluent.Enabled = True
   tmrScroll.Enabled = False
  Else
   tmrFluent.Enabled = False
   tmrScroll.Enabled = True
  End If
  lab(1).Visible = True
 
 Else
   
   tmrFluent.Enabled = False
   tmrScroll.Enabled = False
   lab(1).Visible = False
 
 End If
 
 
 lab(0).ForeColor = GlbWndColor
 lab(1).ForeColor = GlbWndColor
End Sub

Private Sub Form_Activate()
wndTitle.Top = wndMain.Top - wndTitle.Height
wndTitle.Left = wndMain.Left + ((wndMain.Width - wndTitle.Width) / 2)
wndLS
End Sub

Private Sub Form_Paint()
wndTitle.Top = wndMain.Top - wndTitle.Height
wndTitle.Left = wndMain.Left + ((wndMain.Width - wndTitle.Width) / 2)
End Sub

Private Sub Form_Resize()
Dim X, Y, MX, MY, A, CL

MY = Me.Height / Screen.TwipsPerPixelY
For Y = 0 To MY
   Me.Line (0, Y * Screen.TwipsPerPixelY)-(Me.Width, Y * Screen.TwipsPerPixelY), RGB((100 / MY * Y), (100 / MY * Y), 0)
Next

' Dim a As Integer, CL As Integer
For A = 0 To Me.Height Step 25
  CL = 70 + (40 * Sin(A))
  Me.Line (0, A)-(Me.Width, A), RGB(CL, CL, 0), BF
Next
End Sub

Private Sub Timer2_Timer()
End Sub


Private Sub tmrFluent_Timer()
On Error Resume Next
Dim LX, LY, LBX, LBY

LX = 0: LY = 0

If lab(0).Width > Me.Width Then

 Sacred = Sacred - 70
 
 If Sacred < -lab(0).Width Then Sacred = 0
 
 If lab(1).Caption <> lab(0).Caption Then
    lab(1).Caption = lab(0).Caption
 End If
 
 lab(0).Left = Sacred
 lab(1).Left = Sacred + lab(0).Width
 
Else
 If lab(0).Left <> Fix((Me.Width / 2) - (lab(0).Width / 2)) Then lab(0).Left = Fix((Me.Width / 2) - (lab(0).Width / 2)): lab(1).Left = lab(0).Width: lab(1).Caption = ""
 
End If

End Sub

Private Sub tmrScroll_Timer()

If lab(0).Width > Me.Width Then
   Dim L As String, R As String, M As String
   If lab(0).Left <> 0 Then lab(0).Left = 0
   R = Right(lab(0).Caption, Len(lab(0).Caption) - 1)
   L = Left(lab(0).Caption, 1)
   M = R + L
   lab(0).Caption = M
End If

End Sub


