VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5400
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12525
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
   ScaleHeight     =   5400
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   5160
      TabIndex        =   10
      Top             =   4560
      Width           =   4335
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   9
      Top             =   3840
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   8
      Top             =   3000
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   7
      Top             =   2160
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   6
      Top             =   1320
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4680
      MaxLength       =   2
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   16
      Top             =   3000
      Width           =   75
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   15
      Top             =   2160
      Width           =   75
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   " "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   9840
      TabIndex        =   14
      Top             =   1320
      Width           =   75
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Grade"
      Height          =   375
      Left            =   8280
      TabIndex        =   13
      Top             =   3000
      Width           =   810
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Percentage"
      Height          =   375
      Left            =   7560
      TabIndex        =   12
      Top             =   2160
      Width           =   1500
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   375
      Left            =   8400
      TabIndex        =   11
      Top             =   1320
      Width           =   660
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter marks of subject5"
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   3960
      Width           =   3045
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Enter marks of subject4"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3120
      Width           =   3045
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Enter marks of subject3"
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   2280
      Width           =   3045
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter marks of subject2"
      Height          =   375
      Left            =   840
      TabIndex        =   1
      Top             =   1440
      Width           =   3045
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter marks of subject1"
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   3045
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Val(Text1.Text)
b = Val(Text2.Text)
c = Val(Text3.Text)
d = Val(Text4.Text)
e = Val(Text5.Text)
total = a + b + c + d + e
per = (total / 500) * 100
If per >= 90 Then
    grade = "A+"
ElseIf per >= 75 And per < 90 Then
    grade = "A"
ElseIf per >= 60 And per < 75 Then
    grade = "B"
ElseIf per >= 45 And per < 60 Then
    grade = "C"
Else
    grade = "F"
End If
Label9.Caption = total
Label10.Caption = per
Label11.Caption = grade
End Sub
