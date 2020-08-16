VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Form6"
   ClientHeight    =   4185
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5655
   DrawMode        =   5  'Not Copy Pen
   LinkTopic       =   "Form6"
   ScaleHeight     =   4185
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(10), i, j As Integer
For i = 0 To 9 'Storing data
    a(i) = Val(InputBox("enter the number"))
Next
For i = 0 To 4 'printing their sum as required
    j = 9
    sum = a(i) + a(j - i)
    Print "sum of " & i + 1; " element"; _
    " & " & j - i + 1; " element"; " = " & sum
    j = j - 1
Next
End Sub

Private Sub Command2_Click()
Form5.Show
Form6.Visible = False
End Sub

Private Sub Command3_Click()
End
End Sub
