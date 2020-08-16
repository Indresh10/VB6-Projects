VERSION 5.00
Begin VB.Form Problem 
   Caption         =   "Operation With Matrix"
   ClientHeight    =   7005
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Mul"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   3
      Top             =   4080
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Sub"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   2
      Top             =   3360
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   1
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Enter Data"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4800
      TabIndex        =   0
      Top             =   1920
      Width           =   1455
   End
End
Attribute VB_Name = "Problem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim m(3, 3), n(3, 3), add(3, 3), s(3, 3), p(3, 3) As Integer
Private Sub Command1_Click()
'Entering data in first matrix
For i = 0 To 2
    For j = 0 To 2
        m(i, j) = Val(InputBox("Enter data", "FIRST MATRIX"))
    Next j
Next i
Print "First matrix-"
For i = 0 To 2
    For j = 0 To 2
        Print m(i, j);
    Next j
Print
Next i
'Entering data for Second Matrix
For i = 0 To 2
    For j = 0 To 2
        n(i, j) = Val(InputBox("Enter data", "SECOND MATRIX"))
    Next j
Next i
Print "Second matrix-"
For i = 0 To 2
    For j = 0 To 2
        Print n(i, j);
    Next j
Print
Next i

End Sub

Private Sub Command2_Click()
'calc
For i = 0 To 2
    For j = 0 To 2
        add(i, j) = m(i, j) + n(i, j)
    Next j
Next i
Print "Addition-"
For i = 0 To 2
    For j = 0 To 2
        Print add(i, j);
    Next j
Print
Next i
End Sub

Private Sub Command3_Click()
For i = 0 To 2
    For j = 0 To 2
        s(i, j) = m(i, j) - n(i, j)
    Next j
Next i
Print "Subtraction-"
For i = 0 To 2
    For j = 0 To 2
        Print s(i, j);
    Next j
Print
Next i
End Sub

Private Sub Command4_Click()
For i = 0 To 2
    For j = 0 To 2
        For k = 0 To 2
        p(i, j) = p(i, j) + m(i, k) * n(k, j)
        Next k
    Next j
Next i

Print "Multiplication-"
For i = 0 To 2
    For j = 0 To 2
        Print p(i, j);
    Next j
Print
Next i
End Sub
