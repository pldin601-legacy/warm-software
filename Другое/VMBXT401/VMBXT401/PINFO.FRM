VERSION 5.00
Begin VB.Form pinfo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   810
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3450
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   3450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   120
      Picture         =   "pinfo.frx":0000
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lblSesc 
      Caption         =   "Doing some events"
      Height          =   435
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "pinfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
