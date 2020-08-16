VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Currency conversion"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   10920
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   360
      Top             =   600
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   4846
      MaxLength       =   9
      TabIndex        =   0
      Top             =   1230
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Spell"
      Default         =   -1  'True
      Height          =   1215
      Left            =   3646
      TabIndex        =   2
      Top             =   2310
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Left            =   8446
      MaxLength       =   2
      TabIndex        =   1
      Text            =   "00"
      Top             =   1200
      Width           =   495
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   1193
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3990
      Width           =   8535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter the value"
      Height          =   360
      Left            =   1725
      TabIndex        =   6
      Top             =   1350
      Width           =   2130
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Height          =   360
      Left            =   10680
      TabIndex        =   5
      Top             =   120
      Width           =   105
   End
   Begin VB.Label Label4 
      Caption         =   "="
      Height          =   495
      Left            =   8086
      TabIndex        =   4
      Top             =   1350
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
no = Text1.Text
dec = Text2.Text
ne = "Rs." + conwords((no)) + " &" + conwords((dec)) + " Paisa Only"
Text3.Text = ne
End Sub

Private Sub Form_Load()
Label3.Caption = Format(Now, "dd mmmm, yyyy hh:mm:ss am/pm")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = 8
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
End If
End Sub



Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 8 Then
    KeyAscii = 8
ElseIf KeyAscii < 48 Or KeyAscii > 57 Then
KeyAscii = 0
End If
End Sub



Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
Label3.Caption = Format(Now, "dd mmmm, yyyy hh:mm:ss am/pm")
End Sub
