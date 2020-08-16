VERSION 5.00
Begin VB.Form procedure 
   Caption         =   "procedure"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   4140
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Sum"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   1800
      Width           =   1215
   End
End
Attribute VB_Name = "procedure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call sum
End Sub
Private Sub sum()
a = Val(InputBox("Enter First No."))
b = Val(InputBox("Enter Second No."))
total = a + b
Print "Total of" & a & "&" & b & "=" & total
End Sub
