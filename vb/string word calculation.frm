VERSION 5.00
Begin VB.Form str_wrd_calc 
   Caption         =   "Form1"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9405
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
   ScaleHeight     =   5760
   ScaleWidth      =   9405
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "String calculation"
      Height          =   855
      Left            =   4080
      TabIndex        =   0
      Top             =   2640
      Width           =   1815
   End
End
Attribute VB_Name = "str_wrd_calc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = InputBox("Enter any string")
l = Len(n)
blanks = 0
For i = 1 To l
    ch = Mid$(n, i, 1)
    If ch = " " Then blanks = blanks + 1
Next
Print "Length:" & l
Print "No. of blanks spaces:" & blanks
Print "No. of characters without spaces:" & l - blanks
Print "No of words:" & blanks + 1
End Sub
