VERSION 5.00
Begin VB.Form frmfor_pass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Forget Password"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1560
      TabIndex        =   1
      Text            =   " "
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Select account type"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   3180
   End
End
Attribute VB_Name = "frmfor_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As ADODB.Recordset
Dim t As New Database1
Private Sub Command1_Click()
Dim m As Long
If Combo1.Text = "Admin" Then
    user = InputBox("Enter User Name", "Forget Password")
    m = InputBox("Now enter The backup code given by developer", "Forget password")
ElseIf Combo1.Text = "Voter" Then
    MsgBox "Please Contact The administrator for Password", vbOKOnly + vbInformation, "Forget Password"
    Exit Sub
Else
    MsgBox "Please select type first", vbCritical, "Forget password"
    Exit Sub
End If
t.Database ("SELECT * FROM Backup where Code=" & Trim(m))
If rs.EOF Then
    MsgBox "Please check the backup code" + vbCr + Chr$(9) + "or" + vbCr + "contact the developer", vbExclamation
Else
    rs1.Open ("SELECT * FROM Admin WHERE User='" + Trim(user) + "'"), t.db, adOpenKeyset, adLockOptimistic
    If rs1.EOF Then
        MsgBox "Sorry user name is incorrect" + vbCr + "Please Try Again", vbCritical, "Forget Password"
    Else
        MsgBox "Login is Succesfull" + vbCr + "Please change the password", vbInformation
        Login.Text1.Text = rs1.Fields(2)
        Login.Combo1.ListIndex = rs1.Fields(4)
        Login.Hide
        MDIForm1.Show
        frmchng_pass.Show
        frmchng_pass.Text1 = rs1.Fields(3)
        frmchng_pass.Text1.Visible = False
        frmchng_pass.Label1.Visible = False
    End If
End If
End Sub

Private Sub Form_Load()
Combo1.AddItem "Admin"
Combo1.AddItem "Voter"
End Sub

Private Sub Form_Unload(Cancel As Integer)
Login.Show
End Sub
