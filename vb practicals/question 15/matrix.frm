VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "matrix"
   ClientHeight    =   6885
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   10785
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu entry 
      Caption         =   "Enter data"
      Begin VB.Menu Matrix 
         Caption         =   "First Matrix"
         Index           =   0
      End
      Begin VB.Menu Matrix 
         Caption         =   "Second Matrix"
         Index           =   1
      End
   End
   Begin VB.Menu opr 
      Caption         =   "Operation"
      Begin VB.Menu op 
         Caption         =   "Addition"
         Index           =   0
      End
      Begin VB.Menu op 
         Caption         =   "Subtraction"
         Index           =   1
      End
      Begin VB.Menu op 
         Caption         =   "Multiplication"
         Index           =   2
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n(3, 3), m(3, 3), a(3, 3), s(3, 3), p(3, 3) As Integer

Private Sub Form_Load()
pp = MsgBox("To enable operation Enter data of second matrix", vbOKOnly + vbInformation, "Important")
End Sub

Private Sub Matrix_Click(Index As Integer)
Select Case Index
    Case 0
        For i = 0 To 2
            For j = 0 To 2
                n(i, j) = Val(InputBox("Enter Values", "First Matrix"))
            Next j
        Next i
        Print "First matrix"
        For i = 0 To 2
            For j = 0 To 2
                Print n(i, j);
            Next j
        Print
        Next i
    Case 1
        For i = 0 To 2
            For j = 0 To 2
                m(i, j) = Val(InputBox("Enter Values", "Second Matrix"))
            Next j
        Next i
        Print "Second Matrix-"
        For i = 0 To 2
            For j = 0 To 2
                Print m(i, j);
            Next j
        Print
        Next i
        opr.Enabled = True
End Select
End Sub

Private Sub op_Click(Index As Integer)
Select Case Index
    Case 0
        For i = 0 To 2
            For j = 0 To 2
                a(i, j) = n(i, j) + m(i, j)
            Next j
        Next i
        Print "Addition-"
        For i = 0 To 2
            For j = 0 To 2
                Print a(i, j);
            Next j
        Print
        Next i
    Case 1
        For i = 0 To 2
            For j = 0 To 2
                s(i, j) = n(i, j) - m(i, j)
            Next j
        Next i
        Print "Subtraction-"
        For i = 0 To 2
            For j = 0 To 2
                Print s(i, j);
            Next j
        Print
        Next i
    Case 2
        For i = 0 To 2
            For j = 0 To 2
                p(i, j) = 0
                For k = 0 To 2
                    p(i, j) = p(i, j) + n(i, k) * m(k, j)
                Next k
                Print p(i, j);
            Next j
        Print
        Next i
        Print "Multiplication-"
        For i = 0 To 2
           For j = 0 To 2
              Print p(i, j);
           Next j
        Print
        Next i
End Select
End Sub
