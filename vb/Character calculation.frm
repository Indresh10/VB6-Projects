VERSION 5.00
Begin VB.Form character 
   Caption         =   "Form1"
   ClientHeight    =   6990
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
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
   ScaleHeight     =   6990
   ScaleWidth      =   12030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Character calculation"
      Height          =   975
      Left            =   5400
      TabIndex        =   0
      Top             =   3240
      Width           =   2055
   End
End
Attribute VB_Name = "character"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
d = 0: v = 0: s = 0: p = 0: b = 0: c = 0
n = InputBox("Enter the string")
l = Len(n)
For i = 1 To l
    ch = Mid$(n, i, 1)
    Select Case ch
        Case 0 To 9
            d = d + 1
        Case "a", "e", "i", "o", "u", "A", "I", "O", "E", "U"
            v = v + 1
        Case "@", "#", "$", "%", "^", "&", "*"
            s = s + 1
        Case ",", ".", ";", ":", "!"
            p = p + 1
        Case " "
            b = b + 1
        Case Else
            c = c + 1
    End Select
Next
Print "Enter String:" & n
Print "Vowels=" & v
Print "Consonent=" & c
Print "Symbols=" & s
Print "Puntuation marks=" & p
Print "Blanks=" & b
Print "Digits=" & d
End Sub
