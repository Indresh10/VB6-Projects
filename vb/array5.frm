VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "5th Question"
   ClientHeight    =   4575
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5835
   LinkTopic       =   "Form5"
   Picture         =   "array5.frx":0000
   ScaleHeight     =   4575
   ScaleWidth      =   5835
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
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim a(10), i, j, sumodd, sumeven As Integer 'Declaring variable
For i = 0 To 9 'Storing data
    a(i) = Val(InputBox("Enter The Number"))
Next: Print "the odd numbers are-": sumodd = 0
For i = 0 To 9 'Printing odd no. & storing their sum
    If a(i) Mod 2 <> 0 Then
    Print a(i)
    sumodd = sumodd + a(i)
    End If
Next
Print "total=" & sumodd 'print their sum
Print "the even numbers are-"
sumeven = 0
For i = 0 To 9 'printing even no. & storing their sum
    If a(i) Mod 2 = 0 Then
    Print a(i)
    sumeven = sumeven + a(i)
    End If
Next
Print "total=" & sumeven 'Printing their sum
End Sub

Private Sub Command2_Click()
Form5.Visible = False
Form4.Show
End Sub

Private Sub Command3_Click()
End
End Sub

Private Sub Command4_Click()
Form5.Visible = False
Form6.Show
End Sub
