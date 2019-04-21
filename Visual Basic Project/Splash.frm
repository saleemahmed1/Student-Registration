VERSION 5.00
Begin VB.Form Splash 
   Caption         =   "Loading........."
   ClientHeight    =   5160
   ClientLeft      =   6225
   ClientTop       =   3270
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Splash.frx":0000
   LinkTopic       =   "Student Information System"
   Picture         =   "Splash.frx":1222E
   ScaleHeight     =   5160
   ScaleWidth      =   6345
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   5280
      Top             =   4440
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

End Sub

Private Sub Timer1_Timer()
Splash.Hide
frmLogin.Show
Timer1.Enabled = False

End Sub
