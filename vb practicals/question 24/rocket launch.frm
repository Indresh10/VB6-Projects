VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   9480
      Top             =   5280
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Launch"
      Height          =   495
      Left            =   6720
      TabIndex        =   1
      Top             =   9720
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   3375
      Left            =   6240
      Picture         =   "rocket launch.frx":0000
      ScaleHeight     =   3315
      ScaleWidth      =   1875
      TabIndex        =   0
      Top             =   6120
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x

Private Sub Command1_Click()
Timer1.Enabled = True
End Sub

Private Sub Form_Load()
x = Picture1.Top
End Sub

Private Sub Timer1_Timer()
x = x - 500
If x <= 0 Then
Picture1.Visible = False
m = MsgBox("Rocket successfully launched")
End
End If
Picture1.Top = x
End Sub
