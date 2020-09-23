VERSION 5.00
Begin VB.UserControl cButton 
   AutoRedraw      =   -1  'True
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
   ControlContainer=   -1  'True
   ScaleHeight     =   1905
   ScaleWidth      =   3540
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   360
      ScaleHeight     =   65
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   1
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cButton"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "cButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Type cRGB
R As Integer
G As Integer
B As Integer
End Type

Public Enum pAlignment
    [Left Justify] = 0
    [Right Justify] = 1
    [Center] = 2
End Enum

Dim pPressed As Boolean
Dim FocusOn As Boolean
Dim pColorBottom  As OLE_COLOR
Dim pColorTop As OLE_COLOR
Dim pForeColor As OLE_COLOR
Dim pfont As StdFont
Public Event Click()
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Property Get Caption() As Variant
Caption = Label1.Caption
cRefresh
End Property
Public Property Let Caption(ByVal vNewValue As Variant)
Label1.Caption = vNewValue
PropertyChanged "Caption"
cRefresh
End Property
Public Property Get ColorTop() As OLE_COLOR
ColorTop = pColorTop
End Property
Public Property Let ColorTop(ByVal vNewValue As OLE_COLOR)
pColorTop = vNewValue
PropertyChanged "ColorTop"
End Property
Public Property Get ColorBottom() As OLE_COLOR
ColorBottom = pColorBottom
End Property
Public Property Let ColorBottom(ByVal vNewValue As OLE_COLOR)
pColorBottom = vNewValue
PropertyChanged "ColorBottom"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = pForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
pForeColor = vNewValue
PropertyChanged "ForeColor"
End Property
Public Property Get Alignment() As pAlignment
Alignment = Label1.Alignment
End Property
Public Property Let Alignment(ByVal vNewValue As pAlignment)
Label1.Alignment = vNewValue
PropertyChanged "Alignment"
End Property
Private Sub Picture1_GotFocus()
FocusOn = True
cRefresh
End Sub
Private Sub Picture1_LostFocus()
FocusOn = False
cRefresh
End Sub
Private Sub Picture1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 32 Then
    pPressed = True
    cRefresh
End If
RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub Picture1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Or KeyCode = 32 Then
    pPressed = False
    cRefresh
    RaiseEvent Click
End If
RaiseEvent KeyUp(KeyCode, Shift)
End Sub
Private Sub Picture1_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
pPressed = False
cRefresh
RaiseEvent MouseUp(Button, Shift, X, Y)
If Button = 1 Then
RaiseEvent Click
End If
End Sub
Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then
pPressed = True
cRefresh
End If
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_InitProperties()
pForeColor = vbWhite
pColorTop = vbWhite
pColorBottom = vbBlue
cRefresh
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Label1.Caption = PropBag.ReadProperty("Caption", "Command1")
ForeColor = PropBag.ReadProperty("ForeColor", vbRed)
ColorBottom = PropBag.ReadProperty("ColorBottom", vbBlack)
ColorTop = PropBag.ReadProperty("ColorTop", vbRed)
Alignment = PropBag.ReadProperty("Alignment", 2)
cRefresh
End Sub
Private Sub UserControl_Resize()
cRefresh
End Sub
Private Sub UserControl_Show()
cRefresh
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Caption", Label1.Caption, "Command1"
PropBag.WriteProperty "ColorBottom", pColorBottom, vbBlack
PropBag.WriteProperty "ColorTop", pColorTop, vbRed
PropBag.WriteProperty "ForeColor", pForeColor, vbRed
PropBag.WriteProperty "Alignment", Label1.Alignment, 2
cRefresh
End Sub
Private Sub cRefresh()
Dim newRGB As cRGB
Dim uHeight As Long
Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer
Dim i As Integer

R = pColorTop And &HFF
G = (pColorTop \ &H100) And &HFF
B = (pColorTop \ &H10000) And &HFF
R2 = pColorBottom And &HFF
G2 = (pColorBottom \ &H100) And &HFF
B2 = (pColorBottom \ &H10000) And &HFF

Picture1.Left = 0
Picture1.Top = 0
Picture1.Width = UserControl.Width
Picture1.Height = UserControl.Height
'Picture1.ScaleMode = 3
uHeight = (Picture1.Height / Screen.TwipsPerPixelY)
For i = 1 To uHeight
    Picture1.Line (0, i - 1)-((Picture1.Width / Screen.TwipsPerPixelX) - 1, i - 1), RGB(R + ((i / uHeight) * (R2 - R)), G + ((i / uHeight) * (G2 - G)), B + ((i / uHeight) * (B2 - B))), BF
Next

If pPressed = False Then
    Picture1.Line (0, 1)-(0, Picture1.Height / Screen.TwipsPerPixelY), RGB(255, 255, 255), BF
    Picture1.Line (1, 0)-(Picture1.Width / Screen.TwipsPerPixelX, 0), RGB(255, 255, 255), BF
    Picture1.Line (0, (Picture1.Height / Screen.TwipsPerPixelY) - 1)-((Picture1.Width / Screen.TwipsPerPixelX) - 2, (Picture1.Height / Screen.TwipsPerPixelY) - 1), RGB(150, 150, 150), BF
    Picture1.Line ((Picture1.Width / Screen.TwipsPerPixelX) - 1, 0)-((Picture1.Width / Screen.TwipsPerPixelX) - 1, (Picture1.Height / Screen.TwipsPerPixelY) - 2), RGB(150, 150, 150), BF
Else
    Picture1.Line (0, 0)-(1, (Picture1.Height / Screen.TwipsPerPixelY) - 1), RGB(120, 120, 120), BF
    Picture1.Line (1, 0)-(Picture1.Width / Screen.TwipsPerPixelX, 1), RGB(120, 120, 120), BF
    Picture1.Line (0, (Picture1.Height / Screen.TwipsPerPixelY) - 1)-((Picture1.Width / Screen.TwipsPerPixelX) - 2, (Picture1.Height / Screen.TwipsPerPixelY) - 1), RGB(0, 0, 0), BF
    Picture1.Line ((Picture1.Width / Screen.TwipsPerPixelX) - 1, 0)-((Picture1.Width / Screen.TwipsPerPixelX) - 1, (Picture1.Height / Screen.TwipsPerPixelY) - 2), RGB(0, 0, 0), BF
End If
If FocusOn = True Then
Picture1.DrawStyle = 2
Picture1.Line (2, 2)-((Picture1.Width / Screen.TwipsPerPixelX) - 3, (Picture1.Height / Screen.TwipsPerPixelY) - 3), RGB(100, 100, 100), B
Picture1.DrawStyle = 0
End If
DisplayText
Picture1.Refresh
End Sub
Public Sub DisplayText()
Dim divText As Integer
Dim displayedText As String
Dim i As Integer, start As Integer
Picture1.ForeColor = pForeColor
divText = 0

Do
divText = divText + 1
Loop While Picture1.TextWidth(Label1.Caption) / divText > (Picture1.Width / Screen.TwipsPerPixelX)

For i = 1 To divText
start = (CInt(Len(Label1.Caption) / divText) * (i - 1)) + 1
displayedText = Mid(Label1.Caption, start, CInt(Len(Label1.Caption) / divText))
Picture1.CurrentY = ((Picture1.Height / Screen.TwipsPerPixelY) * i / divText) - (((Picture1.Height / Screen.TwipsPerPixelY) / divText) / 2) - (Picture1.TextHeight(Label1.Caption) / 2)

Select Case Label1.Alignment
    Case Is = 0
    Picture1.CurrentX = Picture1.TextWidth(" ")
    Case Is = 1
    Picture1.CurrentX = ((Picture1.Width / Screen.TwipsPerPixelY)) - (Picture1.TextWidth(displayedText & " "))
    Case Is = 2
    Picture1.CurrentX = ((Picture1.Width / Screen.TwipsPerPixelY) / 2) - (Picture1.TextWidth(displayedText) / 2)
End Select

If pPressed = True Then
    Picture1.CurrentX = Picture1.CurrentX + 1
    Picture1.CurrentY = Picture1.CurrentY + 1
End If
Picture1.Print displayedText
Next
End Sub
