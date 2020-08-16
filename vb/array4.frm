VERSION 5.00
Begin VB.Form Form4 
   Caption         =   "4th Question"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6540
   LinkTopic       =   "Form4"
   ScaleHeight     =   4800
   ScaleWidth      =   6540
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(10), i, j As Integer 'declaring variables
For i = 0 To 9 'storing data
    a(i) = Val(InputBox("Enter The Number"))
Next
Print "the odd number are-"
For i = 0 To 9 'printing odd no.
    If a(i) Mod 2 <> 0 Then Print a(i)
Next
Print "the even number are-"
For i = 0 To 9 'printing even no.
    If a(i) Mod 2 = 0 Then Print a(i)
Next
End Sub

Private Sub Command2_Click()
Form3.Show
Form4.Visible = False
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form5.Show
Form4.Visible = False
End Sub
