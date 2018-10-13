VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Редактор Пицц"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Отмена"
      Height          =   375
      Left            =   5040
      TabIndex        =   11
      Top             =   3780
      Width           =   1515
   End
   Begin VB.CommandButton Command4 
      Caption         =   "ОК"
      Height          =   375
      Left            =   6660
      TabIndex        =   10
      Top             =   3780
      Width           =   1515
   End
   Begin VB.Frame Frame3 
      Caption         =   "Навигация"
      Height          =   1035
      Left            =   3300
      TabIndex        =   6
      Top             =   2160
      Width           =   4875
      Begin VB.CommandButton Command3 
         Caption         =   "Изменить"
         Height          =   375
         Left            =   1740
         TabIndex        =   9
         Top             =   360
         Width           =   1395
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Удалить"
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Добавить"
         Height          =   375
         Left            =   180
         TabIndex        =   7
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "В базе:"
      Height          =   3915
      Left            =   180
      TabIndex        =   4
      Top             =   240
      Width           =   2955
      Begin VB.ListBox List1 
         Height          =   3375
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   2595
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Редактор"
      Height          =   1815
      Left            =   3300
      TabIndex        =   0
      Top             =   240
      Width           =   4875
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1260
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   300
         Width           =   3315
      End
      Begin VB.Label Label3 
         Caption         =   "1"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "Название"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
