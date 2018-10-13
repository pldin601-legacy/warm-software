VERSION 5.00
Begin VB.UserControl QSButton 
   Appearance      =   0  'Flat
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   CanGetFocus     =   0   'False
   ClientHeight    =   2460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   2460
   ScaleWidth      =   4800
   Begin VB.Label btnCaption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   360
      Width           =   1260
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
      BorderColor     =   &H00E0E0E0&
      X1              =   180
      X2              =   1500
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   60
      X2              =   60
      Y1              =   0
      Y2              =   780
   End
End
Attribute VB_Name = "QSButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim HX, HY
'Default Property Values:
Const m_def_FontTransparent = 0
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Происходит, когда  пользователь нажимает и затем отпускает кнопку мыши над объектом."






Private Sub btnCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
btnCaption.Left = HX + 15
btnCaption.Top = HY + 15
Line1.BorderColor = RGB(25, 25, 25)
Line2.BorderColor = RGB(25, 25, 25)
Line3.BorderColor = RGB(165, 165, 165)
Line4.BorderColor = RGB(165, 165, 165)
End Sub


Private Sub btnCaption_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then RaiseEvent Click
If Button = 2 Then RaiseEvent MouseDown(Button, Shift, X, Y)
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
If Button = 1 Then RaiseEvent Click
If Button = 2 Then RaiseEvent MouseDown(Button, Shift, X, Y)
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

Line3.X1 = UserControl.Width - 40
Line3.Y1 = 0
Line3.X2 = UserControl.Width - 40
Line3.Y2 = UserControl.Height - 40

Line4.X1 = UserControl.Width - 40
Line4.Y1 = UserControl.Height - 40
Line4.X2 = 0
Line4.Y2 = UserControl.Height - 40

'btnCaption.Left = (UserControl.Width / 2) - (btnCaption.Width / 2)
' btnCaption.Top = (UserControl.Height / 2) - (btnCaption.Height / 2)

btnCaption.Left = 0
btnCaption.Top = (UserControl.Height / 2) - (btnCaption.Height / 2)
btnCaption.Width = UserControl.Width

HX = btnCaption.Left
HY = btnCaption.Top

End Sub


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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,Caption
Public Property Get Caption() As String
Attribute Caption.VB_Description = "Возвращает/Устанавливает текст, отображаемый в области заголовка объекта или ниже иконки объекта."
    Caption = btnCaption.Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
    btnCaption.Caption() = New_Caption
    PropertyChanged "Caption"
End Property

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

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Возвращает объект Шрифт."
Attribute Font.VB_UserMemId = -512
    Set Font = btnCaption.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set btnCaption.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Возвращает/Устанавливает жирный стиль шрифта."
    FontBold = btnCaption.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    btnCaption.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Возвращает/Устанавливает курсивный стиль шрифта."
    FontItalic = btnCaption.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    btnCaption.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Определяет имя шрифта, который появляется в каждой строке для данного уровня."
    FontName = btnCaption.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    btnCaption.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Определяет размер (в пунктах) шрифта, который появляется в каждой строке для данного уровня."
    FontSize = btnCaption.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    btnCaption.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Возвращает/Устанавливает перечеркнутый стиль шрифта."
    FontStrikethru = btnCaption.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    btnCaption.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,3,2,0
Public Property Get FontTransparent() As Boolean
Attribute FontTransparent.VB_Description = "Nothing :-)))"
Attribute FontTransparent.VB_MemberFlags = "400"
    If Ambient.UserMode Then Err.Raise 393
    FontTransparent = m_FontTransparent
End Property

Public Property Let FontTransparent(ByVal New_FontTransparent As Boolean)
    If Ambient.UserMode = False Then Err.Raise 387
    If Ambient.UserMode Then Err.Raise 382
    m_FontTransparent = New_FontTransparent
    PropertyChanged "FontTransparent"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Возвращает/Устанавливает подчеркнутый стиль шрифта."
    FontUnderline = btnCaption.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    btnCaption.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Возвращает/Устанавливает цвет символа, используемый, чтобы отобразить текст и графику в объекте"
    ForeColor = btnCaption.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    btnCaption.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=btnCaption,btnCaption,-1,ToolTipText
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Возвращает/Устанавливает отображаемый текст, когда мышь приостановлена над управлением."
    ToolTipText = btnCaption.ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    btnCaption.ToolTipText() = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property


'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H808080)
    btnCaption.Caption = PropBag.ReadProperty("Caption", "Caption")
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set btnCaption.Font = PropBag.ReadProperty("Font", Ambient.Font)
    btnCaption.FontBold = PropBag.ReadProperty("FontBold", 0)
    btnCaption.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    ' btnCaption.FontName = PropBag.ReadProperty("FontName", "")
    btnCaption.FontSize = PropBag.ReadProperty("FontSize", 0)
    btnCaption.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    m_FontTransparent = PropBag.ReadProperty("FontTransparent", m_def_FontTransparent)
    btnCaption.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    btnCaption.ForeColor = PropBag.ReadProperty("ForeColor", &HFFFFFF)
    btnCaption.ToolTipText = PropBag.ReadProperty("ToolTipText", "")
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H808080)
    Call PropBag.WriteProperty("Caption", btnCaption.Caption, "Caption")
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Font", btnCaption.Font, Ambient.Font)
    Call PropBag.WriteProperty("FontBold", btnCaption.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", btnCaption.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", btnCaption.FontName, "MS Sans Serif")
    Call PropBag.WriteProperty("FontSize", btnCaption.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", btnCaption.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", btnCaption.FontUnderline, 0)
    Call PropBag.WriteProperty("ForeColor", btnCaption.ForeColor, &HFFFFFF)
    Call PropBag.WriteProperty("ToolTipText", btnCaption.ToolTipText, "")
End Sub

