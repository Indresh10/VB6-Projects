VERSION 5.00
Begin VB.Form frmadmin 
   BackColor       =   &H00FFFF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Administrator"
   ClientHeight    =   2385
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4515
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadmin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1409.137
   ScaleMode       =   0  'User
   ScaleWidth      =   4239.34
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPassword 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      IMEMode         =   3  'DISABLE
      Left            =   2145
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   870
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H0080C0FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   2235
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1725
      Width           =   1260
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H0080C0FF&
      Caption         =   "Login"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   750
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1725
      Width           =   1260
   End
   Begin VB.TextBox txtUserName 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   2145
      TabIndex        =   0
      Top             =   120
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   1
      Left            =   240
      TabIndex        =   5
      Top             =   885
      Width           =   1800
   End
   Begin VB.Label lblLabels 
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Top             =   135
      Width           =   1920
   End
End
Attribute VB_Name = "frmadmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
t.Database ("Select * from Admin Where User='" + txtUserName.Text + "'")
If t.rs.EOF Or StrComp(Trim(txtUserName.Text), t.rs.Fields("User"), vbBinaryCompare) <> 0 Or StrComp(Trim(txtPassword.Text), t.rs.Fields("Password"), vbBinaryCompare) <> 0 Then
    MsgBox "Wrong password or username", vbExclamation
    txtUserName.Text = ""
    txtPassword = ""
    txtUserName.SetFocus
Else
    LoginSucceeded = True
    Login.come.Enabled = True
    Unload Me
End If
End Sub
