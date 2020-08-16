VERSION 5.00
Begin VB.Form Login 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   10755
   ClientLeft      =   8235
   ClientTop       =   3210
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Lovely Fabulous"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "Login.frx":10CA
   ScaleHeight     =   10755
   ScaleWidth      =   20250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer delay 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   5880
      Top             =   3120
   End
   Begin VB.Timer go 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   6120
      Top             =   2520
   End
   Begin VB.Timer come 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5760
      Top             =   2400
   End
   Begin VB.CommandButton Command8 
      BackColor       =   &H000000FF&
      Caption         =   "Super Admin"
      Height          =   735
      Left            =   18360
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10845
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0080FF80&
      Caption         =   "Admin Login"
      Default         =   -1  'True
      Height          =   1695
      Left            =   3274
      Picture         =   "Login.frx":5DC9B
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4515
      Width           =   3255
   End
   Begin VB.CommandButton Command11 
      BackColor       =   &H008080FF&
      Caption         =   "Voter Mode"
      Height          =   1695
      Left            =   13721
      Picture         =   "Login.frx":5E3BA
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4545
      Width           =   3255
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Exit"
      Height          =   1695
      Left            =   8494
      Picture         =   "Login.frx":5EAD9
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4515
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Voter Login"
      Height          =   6255
      Left            =   7718
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command9 
         Caption         =   "Login"
         Height          =   510
         Left            =   960
         TabIndex        =   23
         Top             =   5520
         Width           =   1215
      End
      Begin VB.CommandButton Command10 
         Caption         =   "Cancel"
         Height          =   510
         Left            =   2640
         TabIndex        =   22
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Text7 
         Height          =   630
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox Text5 
         Height          =   630
         Left            =   2280
         TabIndex        =   14
         Top             =   2400
         Width           =   2175
      End
      Begin VB.ComboBox Combo4 
         Height          =   630
         Left            =   2280
         TabIndex        =   16
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "PassWord"
         Height          =   495
         Left            =   480
         TabIndex        =   20
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Voter Id"
         Height          =   495
         Left            =   480
         TabIndex        =   19
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   615
         Left            =   480
         TabIndex        =   18
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   1935
         Left            =   1560
         Picture         =   "Login.frx":5F19A
         Stretch         =   -1  'True
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forget Password"
         BeginProperty Font 
            Name            =   "Lovely Fabulous"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   17
         Top             =   4920
         Width           =   1800
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Admin Login"
      Height          =   6255
      Left            =   7680
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   4815
      Begin VB.CommandButton Command2 
         Caption         =   "Cancel"
         Height          =   510
         Left            =   2400
         TabIndex        =   4
         Top             =   5520
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   630
         IMEMode         =   3  'DISABLE
         Left            =   2280
         TabIndex        =   0
         Top             =   2520
         Width           =   2175
      End
      Begin VB.ComboBox Combo1 
         Height          =   630
         Left            =   2280
         TabIndex        =   2
         Top             =   4200
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Height          =   630
         IMEMode         =   3  'DISABLE
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   3360
         Width           =   2175
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Login"
         Height          =   510
         Left            =   840
         TabIndex        =   3
         Top             =   5520
         Width           =   1215
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Forget Password"
         BeginProperty Font 
            Name            =   "Lovely Fabulous"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2880
         TabIndex        =   6
         Top             =   4920
         Width           =   1800
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   1560
         Picture         =   "Login.frx":69896
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   615
         Left            =   480
         TabIndex        =   9
         Top             =   4320
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         Height          =   495
         Left            =   480
         TabIndex        =   8
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "PassWord"
         Height          =   495
         Left            =   480
         TabIndex        =   7
         Top             =   3360
         Width           =   1455
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ra1000"
      ForeColor       =   &H000000FF&
      Height          =   510
      Left            =   13440
      TabIndex        =   28
      Top             =   9400
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Login.frx":73076
      Height          =   1080
      Left            =   4080
      TabIndex        =   27
      Top             =   8880
      Visible         =   0   'False
      Width           =   11475
      WordWrap        =   -1  'True
   End
   Begin VB.Image Image3 
      Height          =   750
      Left            =   9800
      Picture         =   "Login.frx":7310E
      Top             =   1320
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H0080C0FF&
      Caption         =   "Voter mode"
      Height          =   510
      Left            =   9360
      TabIndex        =   25
      Top             =   720
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   1815
      Left            =   9218
      Shape           =   4  'Rounded Rectangle
      Top             =   480
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label lblelection 
      AutoSize        =   -1  'True
      Caption         =   "OFF"
      Height          =   615
      Left            =   10440
      TabIndex        =   26
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label lblvomode 
      AutoSize        =   -1  'True
      Caption         =   "OFF"
      Height          =   615
      Left            =   9600
      TabIndex        =   24
      Top             =   1080
      Visible         =   0   'False
      Width           =   495
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs1 As ADODB.Recordset
Dim t As New Database1
Dim rs2 As ADODB.Recordset
Dim cnt As Integer
Private Function v_id(ByVal name, ByVal Class, ByVal year, ByVal dob, ByVal id) As String
Dim vid As String
vid = "CC10"
vid = vid + Left$(name, 2)
t.Database ("select ID from Class where Name='" + Class + "'")
vid = vid & t.rs.Fields(0)
t.rs.Close
t.db.Close
Select Case year
    Case "First"
        vid = vid + "01"
    Case "Second"
        vid = vid + "02"
    Case "Third"
        vid = vid + "03"
End Select
vid = vid & id
v_id = vid
End Function

Private Function clear()
Text6.Text = ""
Combo3.Text = ""
Text4.Text = ""
Text3.Text = ""
Combo2.Text = ""
DTPicker1.Value = Format(Date, "mm/dd/yy")
Text6.SetFocus
End Function


Private Sub Command1_Click()
Dim user, pass As String
user = Trim(Text1.Text)
pass = Trim(Text2.Text)
If Not Combo1.Text = "" Then
    t.Database ("SELECT * FROM  Admin WHERE User='" & user & "' and Type='" + Combo1.Text + "'")
    If Not t.rs.EOF Then pass1 = t.rs.Fields(3): user1 = t.rs.Fields(2)
Else
    MsgBox "Please Select Type Of Login", vbCritical, "Login"
    Exit Sub
End If
If t.rs.EOF Or StrComp(pass, pass1, vbBinaryCompare) <> 0 Or StrComp(user, user1, vbBinaryCompare) <> 0 Then
    MsgBox "Please Check Your Username, Password And Type", vbCritical, "Login"
    Text1.Text = ""
    Text2.Text = ""
    Combo1.Text = ""
    Text1.SetFocus
Else
    MsgBox "Login Successfull", vbInformation, "Login"
    MDIForm1.Show
    Login.Hide
End If
End Sub

Private Sub Command10_Click()
frmadmin.Show
End Sub

Private Sub Command11_Click()
come.Enabled = True
End Sub

Private Sub Command2_Click()
Frame1.Visible = False
vis_but (True)
Command1.Default = False
End Sub


Private Sub Command3_Click()
vis_but (False)
Frame1.Visible = True
Text1.SetFocus
Command1.Default = True
End Sub

Private Sub Command4_Click()
come.Enabled = True
End Sub

Private Sub Command5_Click()
End
End Sub

Private Sub Command8_Click()
admin_Login.Show
End Sub

Private Sub Command9_Click()
Dim user, pass As String
user = Trim(Text5.Text)
pass = Trim(Text7.Text)
If Not Combo4.Text = "" Then
    t.Database ("SELECT * FROM Voter WHERE V_ID='" & user & "' and Class='" + Combo4.Text + "'")
    If Not t.rs.EOF Then pass1 = t.rs.Fields(6)
Else
    MsgBox "Please Select Your Class", vbCritical, "Login"
    Exit Sub
End If
If t.rs.EOF Or StrComp(pass, pass1, vbBinaryCompare) <> 0 Then
    MsgBox "Please Check Your Username, Password And Class", vbCritical, "Login"
    Text5.Text = ""
    Text7.Text = ""
    Combo4.Text = ""
    Text5.SetFocus
ElseIf t.rs.Fields("Voted") = "No" Then
    MsgBox "Login Successfull", vbInformation, "Login"
    If lblelection.Caption = "OFF" Then
        frmeditdata.Show vbModal, Me
    Else
        frmvote.Show
        Login.Hide
    End If
Else
    MsgBox "Already Voted!!", vbCritical
    Text5.Text = ""
    Text7.Text = ""
    Combo4.Text = ""
    Text5.SetFocus
End If
End Sub

Private Sub delay_Timer()
cnt = cnt + 1
If cnt = 3 Then
        Image3.Picture = LoadPicture(App.Path & "\logo\bet.gif")
ElseIf cnt = 4 Then
    If lblvomode.Caption = "OFF" Then
        Image3.Picture = LoadPicture(App.Path & "\logo\on.gif")
        lblvomode.Caption = "ON"
        vomode (True)
    Else
        Image3.Picture = LoadPicture(App.Path & "\logo\off.gif")
        lblvomode.Caption = "OFF"
        vomode (False)
    End If
End If
If cnt = 5 Then
go.Enabled = True
delay.Enabled = False
cnt = 0
End If
End Sub

Private Sub Form_Load()
t.Database ("Select Name from Courses")
While Not t.rs.EOF
    Combo4.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
t.rs.Close
t.db.Close
t.Database ("select distinct Type from Admin")
While Not t.rs.EOF
    Combo1.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
t.Database ("Select * from status")
lblelection.Caption = t.rs.Fields(0)
lblvomode.Caption = t.rs.Fields(1)
Shape1.Width = 15
Shape1.Height = 15
cnt = 0
lblvomode.Caption = "OFF"
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbBlack
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = vbBlack
End Sub

Private Sub Label11_Click()
frmfor_pass.Show
End Sub

Private Sub Label11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label11.ForeColor = vbRed
End Sub

Private Sub Label4_Click()
frmfor_pass.Show
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.ForeColor = vbRed
End Sub

Private Function vis_but(state As Boolean)
Command3.Visible = state
Command11.Visible = state
Command5.Visible = state
Command8.Visible = state
End Function

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then
    KeyAscii = KeyAscii - 32
End If
End Sub

Private Sub come_Timer()
Shape1.Visible = True
Shape1.Width = Shape1.Width + 100
Shape1.Height = Shape1.Height + 100
Image3.Picture = LoadPicture(App.Path & "\logo\" + LCase(lblvomode.Caption) & ".gif")
If Shape1.Width = 1815 Then
come.Enabled = False
Label5.Visible = True
Image3.Visible = True
delay.Enabled = True
End If
End Sub

Private Sub go_Timer()
Label5.Visible = False
Image3.Visible = False
Shape1.Width = Shape1.Width - 100
Shape1.Height = Shape1.Height - 100
If Shape1.Width = 15 Then
Shape1.Visible = False
go.Enabled = False
End If
End Sub
Private Function vomode(mode As Boolean)
Frame3.Visible = mode
vis_but (Not mode)
Command9.Default = mode
If lblvomode.Caption = "OFF" Then
    Label6.Visible = mode: Label7.Visible = mode
Else
    If lblelection.Caption = "OFF" Then Label6.Visible = mode: Label7.Visible = mode
End If
End Function
