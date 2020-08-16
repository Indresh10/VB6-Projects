VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4395
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10440
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
   ScaleHeight     =   4395
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "remove item from list2"
      Height          =   1215
      Left            =   3960
      TabIndex        =   6
      Top             =   3120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "copy all"
      Height          =   495
      Index           =   3
      Left            =   3960
      TabIndex        =   5
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "copy>>"
      Height          =   495
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   1920
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "move all"
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   3
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "move>>"
      Height          =   495
      Index           =   0
      Left            =   3960
      TabIndex        =   2
      Top             =   720
      Width           =   1815
   End
   Begin VB.ListBox List2 
      Height          =   3435
      Left            =   6960
      TabIndex        =   1
      Top             =   720
      Width           =   2775
   End
   Begin VB.ListBox List1 
      Height          =   3435
      Left            =   600
      TabIndex        =   0
      Top             =   720
      Width           =   2295
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        List2.AddItem List1.Text
        List1.RemoveItem List1.ListIndex
    Case 1
        For i = 0 To List1.ListCount - 1
            List2.AddItem List1.List(i)
        Next
        List1.Clear
    Case 2
        List2.AddItem List1.Text
    Case 3
        For i = 0 To List1.ListCount - 1
            List2.AddItem List1.List(i)
        Next
End Select
End Sub

Private Sub Command2_Click()
List2.RemoveItem List2.ListIndex
End Sub

Private Sub Form_Load()
List1.AddItem "BCA I"
List1.AddItem "BCA II"
List1.AddItem "BCA III"
List2.Clear
End Sub
