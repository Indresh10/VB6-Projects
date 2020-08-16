VERSION 5.00
Begin VB.Form admin_Login 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Super Admin"
   ClientHeight    =   2520
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4425
   Icon            =   "admin_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1488.899
   ScaleMode       =   0  'User
   ScaleWidth      =   4154.835
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2010
      TabIndex        =   1
      Top             =   135
      Width           =   2325
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
      Left            =   735
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1740
      Width           =   1260
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
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1740
      Width           =   1260
   End
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
      Left            =   2010
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   885
      Width           =   2325
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
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1920
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
      Left            =   105
      TabIndex        =   2
      Top             =   900
      Width           =   1800
   End
End
Attribute VB_Name = "admin_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim t As New Database1
Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Me.Hide
End Sub

Private Sub cmdOK_Click()
t.Database ("Select * from Admin Where User='" + txtUserName.Text + "' and Type='Administrator'")
If t.rs.EOF Or StrComp(Trim(txtUserName.Text), t.rs.Fields("User"), vbBinaryCompare) <> 0 Or StrComp(Trim(txtPassword.Text), t.rs.Fields("Password"), vbBinaryCompare) <> 0 Then
    MsgBox "Wrong password or username", vbExclamation
    txtUserName.Text = ""
    txtPassword = ""
    txtUserName.SetFocus
Else
    ad_master.Show
    txtUserName.Text = ""
    txtPassword.Text = ""
    Unload Me
End If
End Sub
