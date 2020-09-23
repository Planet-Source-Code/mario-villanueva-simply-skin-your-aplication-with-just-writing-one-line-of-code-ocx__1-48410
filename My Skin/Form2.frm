VERSION 5.00
Object = "*\A..\MYSKIN~1\Proyecto1.vbp"
Begin VB.Form Form2 
   BackColor       =   &H8000000D&
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   2715
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5055
   LinkTopic       =   "Form2"
   ScaleHeight     =   2715
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.cButton cButton2 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "cButton"
      ColorBottom     =   16711680
      ColorTop        =   16777215
      ForeColor       =   16777215
   End
   Begin Proyecto1.cButton cButton1 
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Caption         =   "cButton"
      ColorBottom     =   16711680
      ColorTop        =   16777215
      ForeColor       =   16777215
   End
   Begin Proyecto1.frmSkin frmSkin1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      Color2          =   16711680
      Color1          =   16777215
      ForeColor       =   16777215
      TitleHeight     =   25
      Degree          =   1
      Style           =   1
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
frmSkin1.SkinForm
End Sub
