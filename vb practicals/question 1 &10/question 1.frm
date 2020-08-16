VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
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
   ScaleHeight     =   5610
   ScaleWidth      =   9720
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Division"
      Height          =   615
      Left            =   7080
      TabIndex        =   9
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Multiplication"
      Height          =   615
      Left            =   4680
      TabIndex        =   8
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Subtraction"
      Height          =   615
      Left            =   2400
      TabIndex        =   7
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Addition"
      Height          =   615
      Left            =   600
      TabIndex        =   6
      Top             =   4320
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Enabled         =   0   'False
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Text            =   " "
      Top             =   3000
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Text            =   " "
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4920
      TabIndex        =   3
      Text            =   " "
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Result"
      Height          =   375
      Left            =   3270
      TabIndex        =   2
      Top             =   3120
      Width           =   810
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter second number"
      Height          =   375
      Left            =   1290
      TabIndex        =   1
      Top             =   1920
      Width           =   2790
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter first number"
      Height          =   375
      Left            =   1815
      TabIndex        =   0
      Top             =   600
      Width           =   2280
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x + y
Text3.Text = z
End Sub

Private Sub Command2_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x - y
Text3.Text = z
End Sub

Private Sub Command3_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x * y
Text3.Text = z
End Sub

Private Sub Command4_Click()
x = Val(Text1.Text)
y = Val(Text2.Text)
z = x / y
Text3.Text = z
End Sub
