VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10815
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
   ScaleHeight     =   5325
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Submit"
      Height          =   855
      Left            =   7800
      TabIndex        =   0
      Top             =   1560
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a(10) As Integer
Private Sub Command1_Click()
For i = 0 To 9
    a(i) = Val(InputBox("Enter the " & i + 1 & " number"))
Next
Print "Entered numbers:"
For i = 0 To 9
    Print a(i)
Next
sk = Val(InputBox("Enter search key"))
beg = 0: cnt = 0: end1 = 9
While beg <= end1
    mid1 = beg + end1 \ 2
    If a(mid1) = sk Then
        cnt = cnt + 1
        GoTo xyz
    End If
    If a(mid1) > sk Then end1 = mid1 - 1
    If a(mid1) < sk Then beg = mid1 + 1
Wend
xyz:
If cnt > 0 Then
    Print "Element found"
Else
    Print "Element not found"
End If
End Sub
