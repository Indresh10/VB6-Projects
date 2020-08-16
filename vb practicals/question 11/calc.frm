VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4380
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   4005
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   525
      Left            =   120
      TabIndex        =   18
      Text            =   " "
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "0"
      Height          =   495
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "1"
      Height          =   840
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2"
      Height          =   855
      Index           =   2
      Left            =   840
      TabIndex        =   15
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "3"
      Height          =   840
      Index           =   3
      Left            =   1560
      TabIndex        =   14
      Top             =   2760
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "4"
      Height          =   855
      Index           =   4
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "5"
      Height          =   855
      Index           =   5
      Left            =   840
      TabIndex        =   12
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "6"
      Height          =   855
      Index           =   6
      Left            =   1560
      TabIndex        =   11
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "7"
      Height          =   735
      Index           =   7
      Left            =   120
      TabIndex        =   10
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "8"
      Height          =   735
      Index           =   8
      Left            =   840
      TabIndex        =   9
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "9"
      Height          =   735
      Index           =   9
      Left            =   1560
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "00"
      Height          =   495
      Index           =   10
      Left            =   1560
      TabIndex        =   7
      Top             =   3720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "+"
      Height          =   735
      Index           =   0
      Left            =   2280
      TabIndex        =   6
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "-"
      Height          =   735
      Index           =   1
      Left            =   3120
      TabIndex        =   5
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "x"
      Height          =   855
      Index           =   2
      Left            =   2280
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "/"
      Height          =   855
      Index           =   3
      Left            =   3120
      TabIndex        =   3
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton Command3 
      Caption         =   "C"
      Height          =   735
      Left            =   2280
      TabIndex        =   2
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "="
      Height          =   735
      Left            =   3120
      TabIndex        =   1
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      Height          =   3615
      Left            =   0
      Top             =   720
      Width           =   3975
   End
   Begin VB.Shape Shape2 
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim val1, val2 As Integer
Dim op
Private Sub Command1_Click(Index As Integer)
Text1.Text = Text1.Text + Command1(Index).Caption
End Sub

Private Sub Command2_Click(Index As Integer)
res = 0
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
val2 = Val(Text1.Text)
Select Case op
    Case "+"
        res = val1 + val2
    Case "-"
        res = val1 - val2
    Case "x"
        res = val1 * val2
    Case "/"
        res = val1 / val2
End Select
Text1.Text = res
End Sub

Private Sub Command5_Click()
End
End Sub
