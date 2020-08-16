VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6435
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10260
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
   ScaleHeight     =   6435
   ScaleWidth      =   10260
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5280
      TabIndex        =   10
      Text            =   " "
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5280
      TabIndex        =   9
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5280
      TabIndex        =   8
      Top             =   720
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   975
      Left            =   2640
      TabIndex        =   3
      Top             =   3480
      Width           =   3975
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   8400
      TabIndex        =   7
      Top             =   5040
      Width           =   75
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   5040
      Width           =   75
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Compound interest"
      Height          =   375
      Left            =   4920
      TabIndex        =   5
      Top             =   5040
      Width           =   2475
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Simple Interest"
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter Time(in years)"
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   2640
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter Rate"
      Height          =   375
      Left            =   3090
      TabIndex        =   1
      Top             =   1680
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter principal"
      Height          =   375
      Left            =   2730
      TabIndex        =   0
      Top             =   720
      Width           =   1845
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
p = Val(Text1.Text)
r = Val(Text2.Text)
t = Val(Text3.Text)
s = (p * r * t) / 100
Label6.Caption = s
a = p * (1 + r / 100) ^ t
c = a - p
Label7.Caption = c
End Sub
