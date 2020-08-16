VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "2nd Question"
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6930
   LinkTopic       =   "Form2"
   ScaleHeight     =   4815
   ScaleWidth      =   6930
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   495
      Left            =   4200
      TabIndex        =   3
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<<"
      Height          =   495
      Left            =   2400
      TabIndex        =   1
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter data"
      Height          =   495
      Left            =   3000
      TabIndex        =   0
      Top             =   2160
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(10), i, j As Integer
For i = 0 To 9
    a(i) = Val(InputBox("Enter The Number"))
Next
Print "the reverse order is-"
For i = 9 To 0 Step -1
    Print a(i)
Next
End Sub

Private Sub Command2_Click()
Form1.Show
Form2.Visible = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form3.Show
Form2.Visible = False
End Sub
