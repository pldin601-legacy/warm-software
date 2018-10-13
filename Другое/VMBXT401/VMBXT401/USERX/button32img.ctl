VERSION 5.00
Begin VB.UserControl QSImgButton 
   BackColor       =   &H00C0C0C0&
   CanGetFocus     =   0   'False
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2460
   ScaleWidth      =   4800
   Begin VB.Image btnCaption 
      Height          =   330
      Left            =   120
      Picture         =   "button32img.ctx":0000
      Top             =   120
      Width           =   495
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00808080&
      X1              =   180
      X2              =   1380
      Y1              =   840
      Y2              =   840
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00808080&
      X1              =   1680
      X2              =   1680
      Y1              =   120
      Y2              =   780
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   180
      X2              =   1500
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   60
      Y1              =   0
      Y2              =   780
   End
End
Attribute VB_Name = "QSImgButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim HX, HY
'Default Property Values:
'Const m_def_FontTransparent = 0
'Event Declarations:
Event Click() 'MappingInfo=btnCaption,btnCaption,-1,Click
'Event Click() 'MappingInfo=UserControl,UserControl,-1,Click




Private Sub btnCaption_Click()
btnCaption.Left = (UserControl.Width / 2) - (btnCaption.Width / 2)
btnCaption.Top = (UserControl.Height / 2) - (btnCaption.Height / 2)

HX = btnCaption.Left
HY = btnCaption.Top
End Sub

Private Sub btnCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnCaption.Left = HX + 15
btnCaption.Top = HY + 15
Line1.BorderColor = RGB(25, 25, 25)
Line2.BorderColor = RGB(25, 25, 25)
Line3.BorderColor = RGB(165, 165, 165)
Line4.BorderColor = RGB(165, 165, 165)
End Sub
'
'
Private Sub btnCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Click
btnCaption.Left = HX
btnCaption.Top = HY
Line1.BorderColor = RGB(255, 255, 255)
Line2.BorderColor = RGB(255, 255, 255)
Line3.BorderColor = RGB(127, 127, 127)
Line4.BorderColor = RGB(127, 127, 127)
End Sub


Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnCaption.Left = HX + 15
btnCaption.Top = HY + 15
Line1.BorderColor = RGB(25, 25, 25)
Line2.BorderColor = RGB(25, 25, 25)
Line3.BorderColor = RGB(165, 165, 165)
Line4.BorderColor = RGB(165, 165, 165)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent Click
btnCaption.Left = HX
btnCaption.Top = HY
Line1.BorderColor = RGB(255, 255, 255)
Line2.BorderColor = RGB(255, 255, 255)
Line3.BorderColor = RGB(127, 127, 127)
Line4.BorderColor = RGB(127, 127, 127)
End Sub


Private Sub UserControl_Resize()
Line1.X1 = 0
Line1.Y1 = 0
Line1.X2 = 0
Line1.Y2 = UserControl.Height - 15

Line2.X1 = 0
Line2.Y1 = 0
Line2.X2 = UserControl.Width - 15
Line2.Y2 = 0

Line3.X1 = UserControl.Width - 15
Line3.Y1 = 0
Line3.X2 = UserControl.Width - 15
Line3.Y2 = UserControl.Height - 15

Line4.X1 = UserControl.Width - 15
Line4.Y1 = UserControl.Height - 15
Line4.X2 = 0
Line4.Y2 = UserControl.Height - 15

btnCaption.Left = (UserControl.Width / 2) - (btnCaption.Width / 2)
btnCaption.Top = (UserControl.Height / 2) - (btnCaption.Height / 2)

HX = btnCaption.Left
HY = btnCaption.Top

End Sub
'
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=UserControl,UserControl,-1,BackColor
'Public Property Get BackColor() As OLE_COLOR
'    BackColor = UserControl.BackColor
'End Property
'
'Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
'    UserControl.BackColor() = New_BackColor
'    PropertyChanged "BackColor"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,Caption
'Public Property Get Caption() As String
'    Caption = btnCaption.Caption
'End Property
'
'Public Property Let Caption(ByVal New_Caption As String)
'    btnCaption.Caption() = New_Caption
'    PropertyChanged "Caption"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Возвращает/Устанавливает значение, которое определяет, может ли объект отвечать на сгенерированные пользователем события."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,Font
'Public Property Get Font() As Font
'    Set Font = btnCaption.Font
'End Property
'
'Public Property Set Font(ByVal New_Font As Font)
'    Set btnCaption.Font = New_Font
'    PropertyChanged "Font"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,FontBold
'Public Property Get FontBold() As Boolean
'    FontBold = btnCaption.FontBold
'End Property
'
'Public Property Let FontBold(ByVal New_FontBold As Boolean)
'    btnCaption.FontBold() = New_FontBold
'    PropertyChanged "FontBold"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,FontItalic
'Public Property Get FontItalic() As Boolean
'    FontItalic = btnCaption.FontItalic
'End Property
'
'Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
'    btnCaption.FontItalic() = New_FontItalic
'    PropertyChanged "FontItalic"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,FontName
'Public Property Get FontName() As String
'    FontName = btnCaption.FontName
'End Property
'
'Public Property Let FontName(ByVal New_FontName As String)
'    btnCaption.FontName() = New_FontName
'    PropertyChanged "FontName"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,FontSize
'Public Property Get FontSize() As Single
'    FontSize = btnCaption.FontSize
'End Property
'
'Public Property Let FontSize(ByVal New_FontSize As Single)
'    btnCaption.FontSize() = New_FontSize
'    PropertyChanged "FontSize"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,FontStrikethru
'Public Property Get FontStrikethru() As Boolean
'    FontStrikethru = btnCaption.FontStrikethru
'End Property
'
'Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
'    btnCaption.FontStrikethru() = New_FontStrikethru
'    PropertyChanged "FontStrikethru"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MemberInfo=0,3,2,0
'Public Property Get FontTransparent() As Boolean
'    If Ambient.UserMode Then Err.Raise 393
'    FontTransparent = m_FontTransparent
'End Property
'
'Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
'    If Ambient.UserMode = False Then Err.Raise 387
'    If Ambient.UserMode Then Err.Raise 382
'    m_FontTransparent = New_FontTransparent
'    PropertyChanged "FontTransparent"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,FontUnderline
'Public Property Get FontUnderline() As Boolean
'    FontUnderline = btnCaption.FontUnderline
'End Property
'
'Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
'    btnCaption.FontUnderline() = New_FontUnderline
'    PropertyChanged "FontUnderline"
'End Property
'
''WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
''MappingInfo=btnCaption,btnCaption,-1,ForeColor
'Public Property Get ForeColor() As OLE_COLOR
'    ForeColor = btnCaption.ForeColor
'End Property
'
'Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
'    btnCaption.ForeColor() = New_ForeColor
'    PropertyChanged "ForeColor"
'End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Возвращает/Устанавливает отображаемый текст, когда мышь приостановлена над управлением."
    ToolTipText = btnCaption.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    btnCaption.ToolTipText = New_ToolTipText
    QSImgButton.ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
'    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H808080)
    
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    btnCaption.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H808080)
End Sub


Private Sub UserControl_Show()
btnCaption.Left = (UserControl.Width / 2) - (btnCaption.Width / 2)
btnCaption.Top = (UserControl.Height / 2) - (btnCaption.Height / 2)

HX = btnCaption.Left
HY = btnCaption.Top
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

'    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H808080)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("ToolTipText", btnCaption.ToolTipText, "")
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H808080)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,Picture
Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Возвращает/Устанавливает   график, который нужно отобразить в управлении."
    Set Picture = btnCaption.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set btnCaption.Picture = New_Picture
    
    btnCaption.Left = (UserControl.Width / 2) - (btnCaption.Width / 2)
    btnCaption.Top = (UserControl.Height / 2) - (btnCaption.Height / 2)

    HX = btnCaption.Left
    HY = btnCaption.Top

    PropertyChanged "Picture"
    
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Возвращает/Устанавливает фоновый цвет, используемый, чтобы отобразить текст и графику в объекте."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

