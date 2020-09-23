VERSION 5.00
Object = "*\AProyecto1.vbp"
Begin VB.Form Form1 
   BackColor       =   &H00FFC0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2880
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2880
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Proyecto1.frmSkin frmSkin1 
      Height          =   2895
      Left            =   360
      TabIndex        =   3
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   5106
      Color2          =   128
      Color1          =   12632319
      ForeColor       =   16777215
      TitleHeight     =   30
      Degree          =   1
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   960
      Width           =   2655
   End
   Begin Proyecto1.cButton cButton8 
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Ok"
      ColorBottom     =   128
      ColorTop        =   12632319
      ForeColor       =   16777215
   End
   Begin Proyecto1.cButton cButton1 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   2040
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      Caption         =   "Cancel"
      ColorBottom     =   128
      ColorTop        =   12632319
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cButton8_Click()
Form2.Show
End Sub

Private Sub Form_Load()
frmSkin1.SkinForm
End Sub
