VERSION 5.00
Begin VB.Form frmBorder 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000001&
   BorderStyle     =   0  'None
   Caption         =   "Virtual MegaBox XT 4.06"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   ScaleHeight     =   90
   ScaleMode       =   0  'User
   ScaleWidth      =   90
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmBorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Show
 
Dim A As Integer, CL As Integer
  
For A = 0 To Me.Height
  CL = 70 + (40 * Sin(A))
  Me.Line (0, A)-(Me.Width, A), RGB(CL, CL, 0), BF
Next

Me.Line (0, 0)-(Me.Width - 25, 0), RGB(155, 155, 0)
Me.Line (0, Me.Height - 30)-(Me.Width - 25, Me.Height - 30), RGB(55, 55, 0)
Me.Line (Me.Width - 25, 0)-(Me.Width - 25, Me.Height - 25), RGB(55, 55, 0)
Me.Line (0, 0)-(0, Me.Height - 30), RGB(155, 155, 0)

Me.Line (15, 15)-(Me.Width - 15 - 25, 15), RGB(100, 100, 0)
Me.Line (15, Me.Height - 15 - 30)-(Me.Width - 25 - 15, Me.Height - 30 - 15), RGB(30, 30, 0)
Me.Line (Me.Width - 15 - 25, 15)-(Me.Width - 25 - 15, Me.Height - 30 - 15), RGB(30, 30, 0)
Me.Line (15, 15)-(15, Me.Height - 30 - 15), RGB(100, 100, 0)


End Sub


