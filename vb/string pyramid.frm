VERSION 5.00
Begin VB.Form str_pyr 
   Caption         =   "Form1"
   ClientHeight    =   5910
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7350
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
   ScaleHeight     =   5910
   ScaleWidth      =   7350
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "String Pyramid"
      Height          =   735
      Left            =   4320
      TabIndex        =   0
      Top             =   2160
      Width           =   2415
   End
End
Attribute VB_Name = "str_pyr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Cls
n = InputBox("Enter string")
For i = 1 To Len(n)
    Print Left$(n, i)
Next
End Sub
