VERSION 5.00
Begin VB.Form frmadd4nos 
   Caption         =   "Form1"
   ClientHeight    =   7815
   ClientLeft      =   3810
   ClientTop       =   2100
   ClientWidth     =   13410
   BeginProperty Font 
      Name            =   "Blackadder ITC"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7815
   ScaleWidth      =   13410
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdgetsum 
      Caption         =   "Get Sum"
      Height          =   855
      Left            =   6480
      TabIndex        =   9
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox txt4 
      Height          =   510
      Left            =   3960
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox txt3 
      Height          =   510
      Left            =   3960
      TabIndex        =   7
      Top             =   3240
      Width           =   1215
   End
   Begin VB.TextBox txt2 
      Height          =   510
      Left            =   3960
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txt1 
      Height          =   510
      Left            =   3960
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label lblresult 
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   10080
      TabIndex        =   4
      Top             =   1920
      Width           =   1650
   End
   Begin VB.Label lbl4no 
      AutoSize        =   -1  'True
      Caption         =   "Enter fourth no."
      Height          =   615
      Left            =   960
      TabIndex        =   3
      Top             =   3960
      Width           =   2865
   End
   Begin VB.Label lbl3no 
      AutoSize        =   -1  'True
      Caption         =   "Enter third no."
      Height          =   615
      Left            =   1200
      TabIndex        =   2
      Top             =   3240
      Width           =   2625
   End
   Begin VB.Label lbl2no 
      AutoSize        =   -1  'True
      Caption         =   "Enter second no."
      Height          =   615
      Left            =   1080
      TabIndex        =   1
      Top             =   2520
      Width           =   2850
   End
   Begin VB.Label lbl1no 
      AutoSize        =   -1  'True
      Caption         =   "Enter first no."
      Height          =   615
      Left            =   1200
      TabIndex        =   0
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "frmadd4nos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdgetsum_Click()
Dim a, b, c, d As Integer
a = Val(txt1.Text)
b = Val(txt2.Text)
c = Val(txt3.Text)
d = Val(txt4.Text)
lblresult.Caption = a + b + c + d
End Sub

