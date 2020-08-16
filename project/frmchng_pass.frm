VERSION 5.00
Begin VB.Form frmchng_pass 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Change Password"
   ClientHeight    =   3570
   ClientLeft      =   6180
   ClientTop       =   3975
   ClientWidth     =   5460
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Poor Richard"
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
   ScaleHeight     =   3570
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command2 
      BackColor       =   &H008080FF&
      Caption         =   "Cancel"
      Height          =   615
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080FFFF&
      Caption         =   "Submit"
      Default         =   -1  'True
      Height          =   615
      Left            =   720
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   1020
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm new password"
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   2955
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter new password"
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      TabIndex        =   1
      Top             =   1020
      Width           =   2580
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter old password"
      ForeColor       =   &H00FF0000&
      Height          =   360
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2460
   End
End
Attribute VB_Name = "frmchng_pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Private Sub Command1_Click()
old = Text1.Text
new1 = Text2.Text
If old = new1 Then
    MsgBox "New Password can't be same as Old Password"
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    If Text1.Visible = False Then
        Text3.SetFocus
    Else
        Text1.SetFocus
    Exit Sub
    End If
End If
If Login.lblvomode.Caption = "OFF" Then
    t.Database ("Select * from Admin where User='" + Login.Text1.Text + "' and Password='" + Trim(old) + "'")
Else
    t.Database ("Select * from Voter where V_ID='" + Login.Text5.Text + "' and Password='" + Trim(old) + "'")
End If
If t.rs.EOF Then
    MsgBox "Please Check your Password", vbExclamation
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    If Text1.Visible = False Then
        Text2.SetFocus
    Else
        Text1.SetFocus
    End If
Else
    t.rs.Update "Password", new1
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    MsgBox "Password changed Succesfull", vbInformation, "Change Password"
    If Login.lblelection.Caption = "OFF" And Login.lblvomode = "ON" Then frmeditdata.Text1.Text = new1
    Unload Me
End If
End Sub


Private Sub Command2_Click()
Unload Me
End Sub


Private Sub Text3_Validate(Cancel As Boolean)
If Text2.Text <> Text3.Text Then
    Cancel = True
    MsgBox "Please check new password", vbExclamation, "Change Password"
    Text2.Text = ""
    Text3.Text = ""
    Text2.SetFocus
End If
End Sub
