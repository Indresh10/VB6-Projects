VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5250
   ClientLeft      =   5970
   ClientTop       =   3405
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   ScaleHeight     =   5250
   ScaleWidth      =   7575
   Begin VB.CommandButton Command1 
      Caption         =   "Check palindrome"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   15
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2400
      TabIndex        =   0
      Top             =   1800
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = InputBox("Enter any word")
m = StrReverse(n)
If m = n Then
    ms = MsgBox("Word is palindrome")
Else
    ms = MsgBox("Word is not palindrome")
End If
End Sub
