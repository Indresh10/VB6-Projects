VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Currency calculator"
   ClientHeight    =   7305
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12900
   BeginProperty Font 
      Name            =   "Museo Sans For Dell"
      Size            =   14.25
      Charset         =   0
      Weight          =   600
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7305
   ScaleWidth      =   12900
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   735
      Left            =   4200
      TabIndex        =   42
      Top             =   6240
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   6
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   31
      Top             =   5280
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   5
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   30
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   4
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   29
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   3
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   28
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   2
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   27
      Top             =   2160
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   1
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   26
      Top             =   1320
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   450
      Index           =   0
      Left            =   8880
      MaxLength       =   3
      TabIndex        =   25
      Top             =   600
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   6
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   5
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   23
      Top             =   4440
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   4
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   22
      Top             =   3720
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   3
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   21
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   2
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   20
      Top             =   2160
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   1
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   19
      Top             =   1320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Index           =   0
      Left            =   5400
      MaxLength       =   2
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select New Notes"
      Height          =   5895
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.CheckBox Check1 
         Caption         =   "2000"
         Height          =   615
         Index           =   6
         Left            =   240
         TabIndex        =   7
         Top             =   5040
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "500"
         Height          =   495
         Index           =   5
         Left            =   240
         TabIndex        =   6
         Top             =   4365
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "200"
         Height          =   615
         Index           =   4
         Left            =   240
         TabIndex        =   5
         Top             =   3555
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "100"
         Height          =   615
         Index           =   3
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "50"
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   3
         Top             =   1965
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "20"
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   2
         Top             =   1155
         Width           =   2055
      End
      Begin VB.CheckBox Check1 
         Caption         =   "10"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Height          =   330
      Left            =   11475
      TabIndex        =   41
      Top             =   6240
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Total"
      Height          =   330
      Left            =   9360
      TabIndex        =   40
      Top             =   6240
      Width           =   675
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   6
      Left            =   11385
      TabIndex        =   39
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   5
      Left            =   11385
      TabIndex        =   38
      Top             =   4440
      Width           =   195
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   4
      Left            =   11385
      TabIndex        =   37
      Top             =   3720
      Width           =   195
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   3
      Left            =   11385
      TabIndex        =   36
      Top             =   2880
      Width           =   195
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   2
      Left            =   11385
      TabIndex        =   35
      Top             =   2160
      Width           =   195
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   1
      Left            =   11385
      TabIndex        =   34
      Top             =   1320
      Width           =   195
   End
   Begin VB.Label lblcur 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "0"
      Height          =   330
      Index           =   0
      Left            =   11385
      TabIndex        =   33
      Top             =   600
      Width           =   195
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Currency"
      Height          =   330
      Left            =   10920
      TabIndex        =   32
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.2000"
      Height          =   330
      Index           =   6
      Left            =   240
      TabIndex        =   17
      Top             =   5400
      Width           =   1110
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.500"
      Height          =   330
      Index           =   5
      Left            =   240
      TabIndex        =   16
      Top             =   4620
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.200"
      Height          =   330
      Index           =   4
      Left            =   240
      TabIndex        =   15
      Top             =   3840
      Width           =   930
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.100"
      Height          =   330
      Index           =   3
      Left            =   240
      TabIndex        =   14
      Top             =   3060
      Width           =   915
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.50"
      Height          =   330
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   2280
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.20"
      Height          =   330
      Index           =   1
      Left            =   240
      TabIndex        =   12
      Top             =   1500
      Width           =   750
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Rs.10"
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Notes"
      Height          =   330
      Index           =   0
      Left            =   240
      TabIndex        =   10
      Top             =   120
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Total Notes"
      Height          =   330
      Left            =   8880
      TabIndex        =   9
      Top             =   120
      Width           =   1545
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Enter serial no.(last 2 di.)"
      Height          =   330
      Left            =   5040
      TabIndex        =   8
      Top             =   120
      Width           =   3300
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l(7)
Private Sub Check1_click(Index As Integer)
If Check1(Index).Value = 1 Then
    Text1(Index).Visible = True
    Text2(Index).Enabled = False
Else
    Text1(Index).Visible = False
    Text2(Index).Enabled = True
End If
End Sub


Private Sub Command1_Click()
f = 0
For i = 0 To 6
l(i) = lblcur(i).Caption
Next
For i = 0 To 6
f = f + l(i)
Next
Label7.Caption = f
End Sub



Private Sub Text1_Change(Index As Integer)
If Text1(Index).Text <> "" Then
    n = Text1(Index).Text
    Text2(Index).Text = (100 - n) + 1
Else
    Text2(Index).Text = ""
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = 8
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub

Private Sub Text2_Change(Index As Integer)
c = Check1(Index).Caption
t = Text2(Index).Text
If Text2(Index).Text <> "" Then lblcur(Index).Caption = c * t
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = 8
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub
