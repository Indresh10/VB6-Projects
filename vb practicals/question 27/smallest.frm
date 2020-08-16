VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   4995
   ClientTop       =   2430
   ClientWidth     =   10335
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
   ScaleHeight     =   5565
   ScaleWidth      =   10335
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   735
      Left            =   3960
      TabIndex        =   6
      Top             =   3360
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Height          =   615
      Left            =   5640
      TabIndex        =   5
      Top             =   2400
      Width           =   2655
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   5640
      TabIndex        =   4
      Top             =   1560
      Width           =   2655
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   6720
      TabIndex        =   8
      Top             =   4560
      Width           =   75
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Smallest Number"
      Height          =   375
      Left            =   3000
      TabIndex        =   7
      Top             =   4680
      Width           =   2235
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter Third number"
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter Second number"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Enter First number"
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   840
      Width           =   2415
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
If a > b Then
    If c > b Then
        s = b
    Else
        s = c
    End If
Else
    If c > a Then
        s = a
    Else
        s = c
    End If
End If
Label5.Caption = s
End Sub
