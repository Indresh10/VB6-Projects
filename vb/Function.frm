VERSION 5.00
Begin VB.Form function 
   Caption         =   "Wap to enter any two no. and print their sum(using funtion)"
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   9330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3840
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
End
Attribute VB_Name = "function"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
x = sumnanr()
Y = sumnawr()
Print "Total:" & Y
n = Val(InputBox("Enter First no"))
m = Val(InputBox("Enter second no"))
z = sumwawr((n), (m))
Print "Total:" & z
a = sumwanr((n), (m))
End Sub
Private Function sumnanr()
n = Val(InputBox("Enter First no"))
m = Val(InputBox("Enter second no"))
total = m + n
Print "Total:" & total
End Function
Private Function sumnawr()
n = Val(InputBox("Enter First no"))
m = Val(InputBox("Enter second no"))
total = m + n
sumnawr = total
End Function
Private Function sumwawr(ByVal x As Integer, ByVal Y As Integer)
total = x + Y
sumwawr = total
End Function
Private Function sumwanr(ByVal x As Integer, ByVal Y As Integer)
total = x + Y
Print "Total:" & total
End Function

