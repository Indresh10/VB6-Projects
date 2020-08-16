VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmadd_cand 
   BackColor       =   &H00FF80FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Candidate"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7380
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadd_cand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      Height          =   525
      Index           =   1
      Left            =   4320
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   4080
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   525
      Index           =   1
      Left            =   4320
      TabIndex        =   9
      Top             =   1080
      Width           =   2775
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6135
      Left            =   120
      ScaleHeight     =   6105
      ScaleWidth      =   7065
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   0
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2880
         Width           =   2775
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   0
         Left            =   4200
         TabIndex        =   8
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancel "
         Height          =   975
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4920
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Save"
         Height          =   975
         Left            =   720
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   4920
         Width           =   2295
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2655
         Left            =   120
         Stretch         =   -1  'True
         Top             =   1320
         Width           =   2295
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Post"
         Height          =   405
         Left            =   2520
         TabIndex        =   5
         Top             =   3960
         Width           =   705
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Party"
         Height          =   405
         Left            =   2520
         TabIndex        =   4
         Top             =   2955
         Width           =   1005
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voter ID "
         Height          =   405
         Left            =   2520
         TabIndex        =   3
         Top             =   1125
         Width           =   1605
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   405
         Left            =   2520
         TabIndex        =   2
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Add Candidate"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   6855
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   3120
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open"
      Filter          =   "*.JPG|JPEG Files(*.jpg)"
   End
End
Attribute VB_Name = "frmadd_cand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim file, fn As String
Dim t As New Database1
Dim t2 As New Database1


Private Sub Command2_Click()
If Image1.Picture = none Or Text1(0).Text = "" Or Text1(1).Text = "" Or Combo1(0).Text = "" Or Combo1(1).Text = "" Then
MsgBox "Check The Field Details", vbExclamation
Exit Sub
End If
t.Database ("Select * from Candidate")
If t.rs.BOF And t.rs.EOF Then
    id = 1
Else
    t.rs.MoveLast
    id = t.rs.Fields(0) + 1
End If
t2.Database ("Select Class,Year,file from Voter Where V_ID='" + Text1(1).Text + "'")
With t.rs
    .AddNew
    .Fields(0) = id
    .Fields(1) = Text1(1).Text
    .Fields(2) = Text1(0).Text
    .Fields(3) = t2.rs.Fields(0)
    .Fields(4) = t2.rs.Fields(1)
    .Fields(5) = t2.rs.Fields(2)
    .Fields(6) = Combo1(0).Text
    .Fields(7) = Combo1(1).Text
    .Update
    .Close
End With
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
t.Database ("Select * from Party")
While Not t.rs.EOF
Combo1(0).AddItem t.rs.Fields(1)
t.rs.MoveNext
Wend
t.rs.Close

t.Database ("Select * from Post")
While Not t.rs.EOF
Combo1(1).AddItem t.rs.Fields(1)
t.rs.MoveNext
Wend
t.rs.Close
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
Select Case Index
    Case 0
        KeyAscii = 0
    Case 1
        If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
    End Select
End Sub

Private Sub Text1_Validate(Index As Integer, Cancel As Boolean)
t.Database ("Select * from Voter where V_ID='" + Text1(1).Text + "'")
If t.rs.EOF Then
    MsgBox "First Register candidate as Voter", vbCritical
    Text1(1).Text = ""
    Text1(1).SetFocus
Else
    If IsNull(t.rs.Fields("file")) Then
    Image1.Picture = LoadPicture(App.Path & "\vo_img\base.gif")
    Else
    Image1.Picture = LoadPicture(App.Path & "\vo_img\" & t.rs.Fields("file"))
    End If
    Text1(0).Text = t.rs.Fields("Name")
End If
End Sub
