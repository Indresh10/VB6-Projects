VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   7770
   ClientTop       =   4200
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Check prime"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   960
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
cnt = 0
n = Val(InputBox("Enter any positive no."))
If n = 1 Then
    m = MsgBox("1 is not prime neither composite")
    GoTo abc
End If
For i = 2 To n - 1
    If n Mod i = 0 Then
        cnt = cnt + 1
        GoTo xyz
    End If
Next
xyz:
If cnt > 0 Then
    m = MsgBox(n & " is not prime")
Else
    m = MsgBox(n & " is prime")
End If
abc:
End Sub
