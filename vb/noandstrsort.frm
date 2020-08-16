VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4920
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton C 
      Caption         =   "clear"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "String"
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Number"
      Height          =   495
      Left            =   3240
      TabIndex        =   0
      Top             =   2280
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n()
Dim d1, d2 As Integer

Private Sub C_Click()
Form1.Cls
End Sub

Private Sub Command1_Click()
d1 = Val(InputBox("enter no. of elements", "numeric expression"))
ReDim n(d1)
ub = d1 - 1
For i = 0 To ub
    n(i) = Val(InputBox("enter the no."))
Next
Print "the entered elements before sort are-"
For i = 0 To ub
    Print n(i)
Next
'sorting
For i = 0 To ub
    For j = 0 To ub - 1
        If n(j) > n(j + 1) Then
        temp = n(j)
        n(j) = n(j + 1)
        n(j + 1) = temp
        End If
    Next j
Next i
Print "elements after sort-"
For i = 0 To ub
    Print n(i)
Next
End Sub

Private Sub Command2_Click()
d2 = Val(InputBox("enter no. of elements", "String expression"))
m = d2 + d1
ReDim n(d1 To m)
ub = (d1 + d2) - 1
For i = 0 To ub
    n(i) = InputBox("enter the strings")
Next
Print "the entered elements before sort are-"
For i = 0 To ub
    Print n(i)
Next
'sorting
For i = 0 To ub
    For j = 0 To ub - 1
        If StrComp(n(j), n(j + 1)) > 0 Then
        temp = n(j)
        n(j) = n(j + 1)
        n(j + 1) = temp
        End If
    Next j
Next i
Print "elements after sort-"
For i = 0 To ub
    Print n(i)
Next
End Sub

