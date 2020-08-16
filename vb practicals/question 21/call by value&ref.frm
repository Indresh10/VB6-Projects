VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4230
   ClientLeft      =   6765
   ClientTop       =   3915
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4230
   ScaleWidth      =   6915
   Begin VB.CommandButton Command1 
      Caption         =   "Swaping"
      Height          =   735
      Left            =   2880
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
n = 10: m = 20
Print "Before swap n = " & n & " m = " & m
Print "call by value"
y = swapv((n), (m)) 'calling by value
Print "After swap: n = " & n & " m = " & m
Print "call by refrence"
x = swapr(n, m) 'call by ref
Print "After swap: n = " & n & " m = " & m
End Sub
Private Function swapv(ByVal a As Integer, ByVal b As Integer)
temp = a
a = b
b = temp
Print " during swap n = " & a & " m = " & b
End Function
Private Function swapr(ByRef a, ByRef b)
temp = a
a = b
b = temp
Print " during swap n = " & a & " m = " & b
End Function

