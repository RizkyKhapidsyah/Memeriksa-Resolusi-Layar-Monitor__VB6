VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Memeriksa Resolusi Layar Monitor"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   6690
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Periksa"
      Height          =   495
      Left            =   5160
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  ResWidth = Screen.Width \ Screen.TwipsPerPixelX
  ResHeight = Screen.Height \ Screen.TwipsPerPixelY
  ScreenRes = ResWidth & "x" & ResHeight
  MsgBox ("Resolusi Layar Monitor Anda Adalah  : " + ScreenRes)
End Sub

