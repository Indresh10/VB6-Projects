VERSION 5.00
Begin VB.Form fact 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7695
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
   ScaleHeight     =   4380
   ScaleWidth      =   7695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Factorial"
      Height          =   855
      Left            =   3240
      TabIndex        =   0
      Top             =   1920
      Width           =   1575
   End
End
Attribute VB_Name = "fact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = Val(InputBox("Enter No."))
p = fact((n))
Print "the factorial of " & n & " is " & p
End Sub
Private Function fact(ByVal a As Integer)
f = 1
If a > 1 Then f = a * fact(a - 1)
fact = f
End Function

