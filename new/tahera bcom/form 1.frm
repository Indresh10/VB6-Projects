VERSION 5.00
Begin VB.Form login 
   BackColor       =   &H00008000&
   Caption         =   "Form1"
   ClientHeight    =   8100
   ClientLeft      =   2925
   ClientTop       =   2280
   ClientWidth     =   15255
   ForeColor       =   &H80000005&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   15255
   Begin VB.CommandButton TA 
      Caption         =   "log in"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   6
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   5
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Height          =   615
      IMEMode         =   3  'DISABLE
      Left            =   5880
      PasswordChar    =   "."
      TabIndex        =   4
      Top             =   3360
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   615
      Left            =   5880
      TabIndex        =   3
      Top             =   2160
      Width           =   2295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "PASSWORD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   2295
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "USER ID"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "welcome to fresh fruit zone "
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   480
      Width           =   9135
   End
End
Attribute VB_Name = "login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub TA_Click()
If Text1.Text = "TAHERA" And Text2.Text = "SHAIKH" Then
home.Show
Unload Me
Else
MsgBox "INVALID DATA"
End If
End Sub
