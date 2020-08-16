VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Calculator"
   ClientHeight    =   4530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   14.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4530
   ScaleWidth      =   5055
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Sqrt"
      Height          =   495
      Index           =   3
      Left            =   3240
      TabIndex        =   23
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "x^2"
      Height          =   495
      Index           =   2
      Left            =   2400
      TabIndex        =   22
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "1/x"
      Height          =   855
      Index           =   1
      Left            =   3240
      TabIndex        =   21
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command7 
      Caption         =   "%"
      Height          =   855
      Index           =   0
      Left            =   2400
      TabIndex        =   20
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   4080
      TabIndex        =   19
      Top             =   3840
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "="
      Height          =   855
      Left            =   4080
      TabIndex        =   18
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Bksp"
      Height          =   855
      Left            =   4080
      TabIndex        =   17
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   735
      Left            =   4080
      TabIndex        =   16
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "/"
      Height          =   855
      Index           =   3
      Left            =   3240
      TabIndex        =   15
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "x"
      Height          =   855
      Index           =   2
      Left            =   2400
      TabIndex        =   14
      Top             =   1920
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   735
      Index           =   1
      Left            =   3240
      TabIndex        =   13
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   735
      Index           =   0
      Left            =   2400
      TabIndex        =   12
      Top             =   960
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "00"
      Height          =   495
      Index           =   10
      Left            =   1680
      TabIndex        =   11
      Top             =   3840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   735
      Index           =   9
      Left            =   1680
      TabIndex        =   10
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   735
      Index           =   8
      Left            =   960
      TabIndex        =   9
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   735
      Index           =   7
      Left            =   240
      TabIndex        =   8
      Top             =   960
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   855
      Index           =   6
      Left            =   1680
      TabIndex        =   7
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   855
      Index           =   5
      Left            =   960
      TabIndex        =   6
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   855
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   840
      Index           =   3
      Left            =   1680
      TabIndex        =   4
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   855
      Index           =   2
      Left            =   960
      TabIndex        =   3
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   840
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   240
      TabIndex        =   0
      Text            =   " "
      Top             =   240
      Width           =   4575
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   120
      Top             =   120
      Width           =   4815
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   120
      Top             =   840
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim val1, val2 As Single
Dim op, exp
Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text + Command1(Index).Caption
End Sub

Private Sub Command2_Click(Index As Integer)
val1 = Val(Text1.Text)
op = Command2(Index).Caption
Text1.Text = ""
End Sub

Private Sub Command3_Click()
Text1.Text = ""
val1 = 0
val2 = 0
End Sub

Private Sub Command4_Click()
Text1.Text = Text1.Text \ 10
End Sub

Private Sub Command5_Click()
val2 = Val(Text1.Text)
Select Case op
     Case "+": res = val1 + val2
     Case "-": res = val1 - val2
     Case "x": res = val1 * val2
     Case "/": res = val1 / val2
End Select
Text1.Text = res
End Sub

Private Sub Command6_Click()
End
End Sub

Private Sub Command7_Click(Index As Integer)
val1 = Val(Text1.Text)
exp = Command7(Index).Caption
Select Case exp
    Case "%": res = val1 / 100
    Case "1/x": res = 1 / val1
    Case "x^2": res = val1 ^ 2
    Case "Sqrt": res = val1 ^ (1 / 2)
End Select
Text1.Text = res
End Sub
