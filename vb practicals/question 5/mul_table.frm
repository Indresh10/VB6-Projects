VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9645
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   9645
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   855
      Left            =   4920
      TabIndex        =   0
      Top             =   1320
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim x, t, i
x = Val(Text1.Text)
    i = 1
    Do While i < 11
        t = x * i
        Print x & " * " & i & " = " & t
        i = i + 1
    Loop
End Sub
