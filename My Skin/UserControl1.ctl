VERSION 5.00
Begin VB.UserControl frmSkin 
   BackStyle       =   0  'Transparent
   ClientHeight    =   1875
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3300
   ScaleHeight     =   1875
   ScaleWidth      =   3300
   Begin VB.PictureBox statusPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      ScaleHeight     =   375
      ScaleWidth      =   2535
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
   End
   Begin VB.PictureBox titlePic 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   120
      ScaleHeight     =   615
      ScaleWidth      =   3135
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.PictureBox resPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00008000&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2160
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.PictureBox minPic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000D&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   1680
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.PictureBox closePic 
         AutoRedraw      =   -1  'True
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   2640
         ScaleHeight     =   375
         ScaleWidth      =   375
         TabIndex        =   1
         Top             =   120
         Width           =   375
      End
   End
   Begin VB.PictureBox leftPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   5
      Top             =   720
      Width           =   255
   End
   Begin VB.PictureBox rightPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   3000
      ScaleHeight     =   1095
      ScaleWidth      =   255
      TabIndex        =   6
      Top             =   720
      Width           =   255
   End
End
Attribute VB_Name = "frmSkin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type cRGB
R As Integer
G As Integer
B As Integer
End Type

Public Enum Deg
    [Horizontal] = 0
    [Vertical] = 1
End Enum

Public Enum TitleR
    [Normal] = 0
    [Rounded] = 1
End Enum

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type WINDOWPLACEMENT
    Length           As Long
    flags            As Long
    showCmd          As Long
    ptMinPosition    As POINTAPI
    ptMaxPosition    As POINTAPI
    rcNormalPosition As RECT
End Type

Dim pDegree As Deg
Dim pStyle As TitleR
Dim pColor2  As OLE_COLOR
Dim pColor1 As OLE_COLOR
Dim pForeColor As OLE_COLOR
Dim pTitleHeight As Integer
Dim frmHwnd As Long
Dim frmWidth As Long
Dim frmHeight As Long
Public Event ExitApp()

Private Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function SetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Property Get Degree() As Deg
Degree = pDegree
End Property
Public Property Let Degree(ByVal vNewValue As Deg)
pDegree = vNewValue
TitleRefresh
PropertyChanged "Degree"
End Property
Public Property Get Style() As TitleR
Style = pStyle
End Property
Public Property Let Style(ByVal vNewValue As TitleR)
pStyle = vNewValue
TitleRefresh
PropertyChanged "Style"
End Property

Public Property Get TitleHeight() As Integer
TitleHeight = pTitleHeight
End Property
Public Property Let TitleHeight(ByVal vNewValue As Integer)
pTitleHeight = vNewValue
TitleRefresh
PropertyChanged "TitleHeight"
End Property
Public Property Get Color1() As OLE_COLOR
Color1 = pColor1
End Property
Public Property Let Color1(ByVal vNewValue As OLE_COLOR)
pColor1 = vNewValue
TitleRefresh
PropertyChanged "Color1"
End Property
Public Property Get Color2() As OLE_COLOR
Color2 = pColor2
End Property
Public Property Let Color2(ByVal vNewValue As OLE_COLOR)
pColor2 = vNewValue
TitleRefresh
PropertyChanged "Color2"
End Property
Public Property Get ForeColor() As OLE_COLOR
ForeColor = pForeColor
End Property
Public Property Let ForeColor(ByVal vNewValue As OLE_COLOR)
pForeColor = vNewValue
TitleRefresh
PropertyChanged "ForeColor"
End Property
Private Sub leftPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
leftPic.MousePointer = 9
If Button = 1 Then
Const WM_NCLBUTTONDOWN = &HA1
Const HTLEFT = 10
ReleaseCapture
SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTLEFT, ByVal 0&
SkinForm
End If
End Sub
Private Sub minPic_Click()
Const WM_SYSCOMMAND = &H112
Const SC_MINIMIZE = &HF020&
SendMessage frmHwnd, WM_SYSCOMMAND, SC_MINIMIZE, 0
End Sub
Private Sub rightPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
rightPic.MousePointer = 9
If Button = 1 Then
Const WM_NCLBUTTONDOWN = &HA1
Const HTRIGHT = 11
ReleaseCapture
SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTRIGHT, ByVal 0&
SkinForm
End If
End Sub

Private Sub statusPic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Const WM_NCLBUTTONDOWN = &HA1
Const HTBOTTOMRIGHT = 17
Const HTBOTTOM = 15
Const HTBOTTOMLEFT = 16
Select Case x
    Case Is > (frmWidth) - (10 * Screen.TwipsPerPixelX)
        ReleaseCapture
        SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTBOTTOMRIGHT, ByVal 0&
        SkinForm
    Case Is < (10 * Screen.TwipsPerPixelX)
        ReleaseCapture
        SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTBOTTOMLEFT, ByVal 0&
        SkinForm
    Case Else
        ReleaseCapture
        SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTBOTTOM, ByVal 0&
        SkinForm
End Select
TitleRefresh
End Sub
Private Sub statusPic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
statusPic.MousePointer = 0
Select Case x
    Case Is > (frmWidth) - (10 * Screen.TwipsPerPixelX)
        statusPic.MousePointer = 8
    Case Is < (10 * Screen.TwipsPerPixelX)
        statusPic.MousePointer = 6
    Case Else
        statusPic.MousePointer = 7
End Select
End Sub
Private Sub titlePic_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl.MousePointer = 0
End Sub
Private Sub UserControl_AmbientChanged(PropertyName As String)
TitleRefresh
End Sub
Private Sub UserControl_Initialize()
pForeColor = vbWhite
pColor1 = vbWhite
pColor2 = vbBlue
pTitleHeight = 20
TitleRefresh
End Sub
Private Sub titlePic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Const RGN_OR = 2
Const WM_NCLBUTTONDOWN = &HA1
Const HTCAPTION = 2
ReleaseCapture
SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub closePic_Click()
Const WM_CLOSE = &H10
SendMessage frmHwnd, WM_CLOSE, 0, vbNullString
End Sub
Private Sub UserControl_InitProperties()
TitleRefresh
End Sub
Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Const WM_NCLBUTTONDOWN = &HA1
Const HTRIGHT = 11
If x > (frmWidth) - (10 * Screen.TwipsPerPixelX) Then
ReleaseCapture
SendMessage frmHwnd, WM_NCLBUTTONDOWN, HTRIGHT, ByVal 0&
SkinForm frmHwnd
End If
TitleRefresh
End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
UserControl.MousePointer = 0
If x > (frmWidth) - (10 * Screen.TwipsPerPixelX) Then
UserControl.MousePointer = 9
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
ForeColor = PropBag.ReadProperty("ForeColor", vbRed)
Color2 = PropBag.ReadProperty("Color2", vbBlack)
Color1 = PropBag.ReadProperty("Color1", vbRed)
TitleHeight = PropBag.ReadProperty("TitleHeight", 20)
Degree = PropBag.ReadProperty("Degree", 0)
Style = PropBag.ReadProperty("Style", 0)
End Sub
Private Sub UserControl_Resize()
TitleRefresh
End Sub
Private Sub UserControl_Show()
TitleRefresh
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
PropBag.WriteProperty "Color2", pColor2, vbBlack
PropBag.WriteProperty "Color1", pColor1, vbRed
PropBag.WriteProperty "ForeColor", pForeColor, vbRed
PropBag.WriteProperty "TitleHeight", pTitleHeight, 20
PropBag.WriteProperty "Degree", pDegree, 0
PropBag.WriteProperty "Style", pStyle, 0
End Sub
Public Function SkinForm()
Dim FullRegion As Long
Dim LineRegion As Long
Dim LineRegion2 As Long
Dim pos As Integer
Dim startY As Double
Dim frmWP As WINDOWPLACEMENT, CtrlWP As WINDOWPLACEMENT

frmHwnd = GetParent(UserControl.hwnd)
CtrlWP.Length = Len(CtrlWP)
frmWP.Length = Len(frmWP)
GetWindowPlacement frmHwnd, frmWP
GetWindowPlacement UserControl.hwnd, CtrlWP
CtrlWP.rcNormalPosition.Left = 0
CtrlWP.rcNormalPosition.Top = 0
CtrlWP.rcNormalPosition.Right = frmWP.rcNormalPosition.Right - frmWP.rcNormalPosition.Left
CtrlWP.rcNormalPosition.Bottom = frmWP.rcNormalPosition.Bottom - frmWP.rcNormalPosition.Top
SetWindowPlacement UserControl.hwnd, CtrlWP

frmWidth = (frmWP.rcNormalPosition.Right - frmWP.rcNormalPosition.Left) * Screen.TwipsPerPixelX
frmHeight = (frmWP.rcNormalPosition.Bottom - frmWP.rcNormalPosition.Top) * Screen.TwipsPerPixelY


LineRegion2 = CreateRectRgn(0, pTitleHeight, (frmWidth / Screen.TwipsPerPixelX) - 1, pTitleHeight + 1)
If Style = 1 Then
    For pos = 0 To pTitleHeight
        startY = pTitleHeight - (pTitleHeight * Sqr(1 - ((1 - (pos / pTitleHeight)) * (1 - (pos / pTitleHeight)))))
        LineRegion = CreateRectRgn(pos, startY, (frmWidth / Screen.TwipsPerPixelX) - pos, pTitleHeight)
        CombineRgn LineRegion2, LineRegion2, LineRegion, 2
    Next
    
    'LineRegion = CreateRectRgn(0, (frmHeight / Screen.TwipsPerPixelY) - (pTitleHeight), (frmWidth / Screen.TwipsPerPixelX), (frmHeight / Screen.TwipsPerPixelY) - (pTitleHeight / 2))
    'CombineRgn LineRegion2, LineRegion2, LineRegion, 2
    
    'For pos = 0 To pTitleHeight / 2
    '    startY = (pTitleHeight) - (pTitleHeight * Sqr(1 - ((1 - (pos / pTitleHeight)) * (1 - (pos / pTitleHeight)))))
    '    startY = startY / 2
    '    LineRegion = CreateRectRgn(startY, (frmHeight / Screen.TwipsPerPixelY) - pos, (frmWidth / Screen.TwipsPerPixelX) - startY, (frmHeight / Screen.TwipsPerPixelY) - pos - 1)
    '    CombineRgn LineRegion2, LineRegion2, LineRegion, 2
    'Next
Else
LineRegion = CreateRectRgn(0, 0, (frmWidth / Screen.TwipsPerPixelX), pTitleHeight)
CombineRgn LineRegion2, LineRegion2, LineRegion, 2
End If

FullRegion = CreateRectRgn(0, pTitleHeight + 1, (frmWidth / Screen.TwipsPerPixelX), (frmHeight / Screen.TwipsPerPixelY)) ' - pTitleHeight)
CombineRgn FullRegion, LineRegion2, FullRegion, 2
SetWindowRgn frmHwnd, FullRegion, True
TitleRefresh
End Function
Private Sub TitleRefresh()
Dim R As Integer, G As Integer, B As Integer
Dim R2 As Integer, G2 As Integer, B2 As Integer
Dim x As Long, y As Long
Dim max As Long, dWidth As Integer

max = pTitleHeight

R = pColor1 And &HFF
G = (pColor1 \ &H100) And &HFF
B = (pColor1 \ &H10000) And &HFF
R2 = pColor2 And &HFF
G2 = (pColor2 \ &H100) And &HFF
B2 = (pColor2 \ &H10000) And &HFF

titlePic.Left = 0
titlePic.Top = 0
titlePic.Width = UserControl.Width
titlePic.Height = max * Screen.TwipsPerPixelY

statusPic.Left = 0
statusPic.Top = frmHeight - (max * Screen.TwipsPerPixelY / 2)
statusPic.Width = frmWidth
statusPic.Height = max * Screen.TwipsPerPixelY / 2

leftPic.Left = 0
leftPic.Top = 0
leftPic.Width = (max * Screen.TwipsPerPixelX / 2)
leftPic.Height = frmHeight

rightPic.Left = frmWidth - (max * Screen.TwipsPerPixelX / 2)
rightPic.Top = 0
rightPic.Width = (max * Screen.TwipsPerPixelX / 2)
rightPic.Height = frmHeight

If Degree = 0 Then
    dWidth = titlePic.Width / Screen.TwipsPerPixelX
    rightPic.BackColor = pColor2
    leftPic.BackColor = pColor1
    For y = 0 To max
        For x = 0 To dWidth
            SetPixel titlePic.hdc, x, y, RGB(R + ((x / dWidth) * (R2 - R)), G + ((x / dWidth) * (G2 - G)), B + ((x / dWidth) * (B2 - B)))
            SetPixel statusPic.hdc, x, y, RGB(R + ((x / dWidth) * (R2 - R)), G + ((x / dWidth) * (G2 - G)), B + ((x / dWidth) * (B2 - B)))
        Next
    Next
Else
    rightPic.BackColor = pColor2
    leftPic.BackColor = pColor2
    For x = 0 To titlePic.Width / Screen.TwipsPerPixelX
        For y = 0 To max
            SetPixel titlePic.hdc, x, y, RGB(R + ((y / max) * (R2 - R)), G + ((y / max) * (G2 - G)), B + ((y / max) * (B2 - B)))
            SetPixel statusPic.hdc, x, y, RGB(R2 + ((y / max * 2) * (R - R2)), G2 + ((y / max * 2) * (G - G2)), B2 + ((y / max * 2) * (B - B2)))
        Next
    Next
End If

titlePic.FontBold = True
titlePic.ForeColor = pForeColor
titlePic.CurrentX = max * Screen.TwipsPerPixelX
titlePic.CurrentY = (max * Screen.TwipsPerPixelY / 2) - (titlePic.TextHeight(App.Title) / 2)
titlePic.Print App.Title

closePic.Left = titlePic.Width - (1.2 * max * Screen.TwipsPerPixelX)
closePic.Top = (max / 4 * Screen.TwipsPerPixelX)
closePic.Width = max / 2 * Screen.TwipsPerPixelX
closePic.Height = max / 2 * Screen.TwipsPerPixelY
closePic.BackColor = pColor2
closePic.ForeColor = pColor1
closePic.DrawWidth = 2 * max / 25
closePic.Line (0, 0)-(closePic.Width, closePic.Height), , B
closePic.Line (0, 0)-(closePic.Width, closePic.Height)
closePic.Line (closePic.Width, 0)-(0, closePic.Height)

minPic.Left = closePic.Left - (1.5 * max / 2 * Screen.TwipsPerPixelX)
minPic.Top = (max / 4 * Screen.TwipsPerPixelX)
minPic.Width = max / 2 * Screen.TwipsPerPixelX
minPic.Height = max / 2 * Screen.TwipsPerPixelY
minPic.BackColor = pColor2
minPic.ForeColor = pColor1
minPic.DrawWidth = 2 * max / 25
minPic.Line (0, 0)-(minPic.Width, minPic.Height), , B
minPic.Line (0, closePic.Height - (minPic.DrawWidth * Screen.TwipsPerPixelY))-(closePic.Width - 1, closePic.Height - (minPic.DrawWidth * Screen.TwipsPerPixelY))

closePic.Refresh
minPic.Refresh
titlePic.Refresh
closePic.Refresh
statusPic.Refresh
End Sub

