VERSION 5.00
Begin VB.Form Try 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[00:00]"
   ClientHeight    =   30
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2760
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30
   ScaleWidth      =   2760
   StartUpPosition =   3  'Windows Default
   WindowState     =   1  'Minimized
End
Attribute VB_Name = "Try"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
Me.WindowState = 1
End Sub


