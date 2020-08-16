VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "First Question"
   ClientHeight    =   4830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7080
   LinkTopic       =   "Form1"
   ScaleHeight     =   4830
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command2 
      Caption         =   "Goto"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">>"
      Height          =   495
      Left            =   4200
      TabIndex        =   2
      Top             =   2640
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Width           =   1215
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
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(10), i, j As Integer
For i = 0 To 9
    a(i) = Val(InputBox("Enter The Number"))
Next
Print "the result is-"
For i = 0 To 9
    Print a(i)
Next
End Sub

Private Sub Command2_Click()
i = Val(InputBox("Goto Form number", "GOTO"))
Select Case i
Case 1: Form1.Show
Case 2: Form2.Show
Case 3: Form3.Show
Case 4: Form4.Show
Case 5: Form5.Show
Case 6: Form6.Show
End Select
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form2.Show
Form1.Visible = False
End Sub
