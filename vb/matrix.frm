VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10935
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   4920
      TabIndex        =   7
      Top             =   5760
      Width           =   3375
   End
   Begin VB.CommandButton Command7 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   4920
      TabIndex        =   6
      Top             =   5160
      Width           =   3375
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Addition of right diagonal"
      Height          =   855
      Left            =   4920
      TabIndex        =   5
      Top             =   4320
      Width           =   3375
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Addition of left diagonal"
      Height          =   855
      Left            =   4920
      TabIndex        =   4
      Top             =   3480
      Width           =   3375
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Addition of column"
      Height          =   615
      Left            =   4920
      TabIndex        =   3
      Top             =   2880
      Width           =   3375
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Addition of rows"
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   2280
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Transpose of matrix"
      Height          =   615
      Left            =   4920
      TabIndex        =   1
      Top             =   1680
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter Data for matrix"
      Height          =   615
      Left            =   4920
      TabIndex        =   0
      Top             =   1080
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m(3, 3), tr(3), tc(0, 3) As Integer
Private Sub Command1_Click()
For i = 0 To 2
    tc(0, i) = 0
Next
'Entering data & calculating the data as required
For i = 0 To 2
    tr(i) = 0
    For j = 0 To 2
        m(i, j) = Val(InputBox("Enter data of matrix", "Playing with matrix"))
        tr(i) = tr(i) + m(i, j)
        tc(0, j) = tc(0, j) + m(i, j)
    Next j
Next i
'Printing data in matrix form
Print "Entered data-"
For i = 0 To 2
    For j = 0 To 2
        Print m(i, j);
    Next j
    Print
Next i
End Sub

Private Sub Command2_Click()
'printing the transpose of matrix
Print "Transpose is-"
For i = 0 To 2
    For j = 0 To 2
        Print m(j, i);
    Next j
    Print
Next i
End Sub

Private Sub Command3_Click()
'printing matrix with its addition of rows
Print "data with addition of row-"
For i = 0 To 2
    For j = 0 To 2
        Print m(i, j);
    Next j
    Print "=" & tr(i)
Next i
End Sub

Private Sub Command4_Click()
'printing matrix with its addition of column
Print "data with addition of column"
For i = 0 To 2
    For j = 0 To 2
        Print m(i, j);
    Next j
    Print
Next i
Print " =  =  ="
For i = 0 To 2
    Print tc(0, i);
Next
Print
End Sub

Private Sub Command5_Click()
'calculating sum of left diagonal
Sum = 0
For i = 0 To 2
    Sum = Sum + m(i, i)
Next
'printing sum
Print "Addition of left diagonal=" & Sum
End Sub

Private Sub Command6_Click()
'calculating sum of right diagonal
Sum = 0
j = 2
For i = 0 To 2
    Sum = Sum + m(i, j)
    j = j - 1
Next
'printing sum
Print "Addition of right diagonal=" & Sum
End Sub

Private Sub Command7_Click()
'clear the contents printed on form
Form1.Refresh
End Sub

Private Sub Command8_Click()
'exit the program
End
End Sub
