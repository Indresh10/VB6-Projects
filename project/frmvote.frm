VERSION 5.00
Begin VB.Form frmvote 
   Caption         =   "Vote"
   ClientHeight    =   10935
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20250
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmvote.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmvote.frx":10CA
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   7
      Left            =   13200
      TabIndex        =   31
      Text            =   "0"
      Top             =   10200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   6
      Left            =   9360
      TabIndex        =   30
      Text            =   "0"
      Top             =   10200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   5
      Left            =   5040
      TabIndex        =   29
      Text            =   "0"
      Top             =   10200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   4
      Left            =   720
      TabIndex        =   28
      Text            =   "0"
      Top             =   10200
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   3
      Left            =   13200
      TabIndex        =   27
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   2
      Left            =   9360
      TabIndex        =   26
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   1
      Left            =   5040
      TabIndex        =   25
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   675
      Index           =   0
      Left            =   720
      TabIndex        =   24
      Text            =   "0"
      Top             =   6000
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   5
      Left            =   5400
      TabIndex        =   22
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5895
      Left            =   16920
      TabIndex        =   17
      Top             =   4320
      Width           =   3015
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Cast my Vote"
         Height          =   1215
         Index           =   2
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   4320
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Reset"
         Height          =   1215
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2340
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFC0FF&
         Caption         =   "Vote Directly to a Party"
         Height          =   1215
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   360
         Width           =   2295
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFF00&
         BorderWidth     =   5
         Height          =   5895
         Left            =   0
         Top             =   0
         Width           =   3015
      End
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   7
      Left            =   13560
      TabIndex        =   16
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   6
      Left            =   9720
      TabIndex        =   15
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   4
      Left            =   1080
      TabIndex        =   14
      Top             =   10200
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   3
      Left            =   13560
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   2
      Left            =   9720
      TabIndex        =   8
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   1
      Left            =   5400
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.ComboBox Combo1 
      Height          =   675
      Index           =   0
      Left            =   1080
      TabIndex        =   4
      Top             =   6000
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      Picture         =   "frmvote.frx":11EFA
      ScaleHeight     =   2835
      ScaleWidth      =   20475
      TabIndex        =   0
      Top             =   -120
      Width           =   20535
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   3240
         Top             =   2280
      End
      Begin VB.Image Image2 
         Height          =   2880
         Left            =   -10
         Picture         =   "frmvote.frx":288A8
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2880
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Welcome"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   15120
         TabIndex        =   23
         Top             =   1920
         Width           =   2115
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "Christ College Jagdalpur"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1710
         Left            =   2760
         TabIndex        =   2
         Top             =   480
         Width           =   5220
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Welcome to Voting Session"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Left            =   6975
         TabIndex        =   1
         Top             =   720
         Width           =   7035
      End
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   5
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   4
      Left            =   1800
      TabIndex        =   21
      Top             =   6960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   7
      Left            =   13560
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   6
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   4
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   7
      Left            =   14280
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   6
      Left            =   10440
      TabIndex        =   12
      Top             =   6960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   5
      Left            =   6120
      TabIndex        =   11
      Top             =   6960
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   3
      Left            =   13560
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   3
      Left            =   14235
      TabIndex        =   9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   2
      Left            =   9720
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   2
      Left            =   10395
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   1
      Left            =   5400
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   1
      Left            =   6075
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2655
      Index           =   0
      Left            =   1080
      Stretch         =   -1  'True
      Top             =   3360
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   0
      Left            =   1755
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "frmvote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Dim t2 As New Database1
Dim t3 As New Database1
Private Function hide1(ByVal state As Boolean, ByVal Index As Integer)
Label3(Index).Visible = state
Image1(Index).Visible = state
Combo1(Index).Visible = state
End Function

Private Sub Combo1_Change(Index As Integer)
If Not Combo1(Index).Text = "" Then
    t.Database ("select ID,file from Candidate where Name='" + Combo1(Index).Text + "'")
    If Not t.rs.EOF Then
    Image1(Index).Picture = LoadPicture(App.Path & "\vo_img\" & t.rs.Fields(1))
    Text1(Index).Text = t.rs.Fields(0)
    t.rs.Close
    t.db.Close
    End If
Else
    Image1(Index).Picture = LoadPicture(App.Path & "\logo\blank.gif")
    Text1(Index).Text = ""
End If
End Sub

Private Sub Combo1_Click(Index As Integer)
If Not Combo1(Index).Text = "" Then
    t.Database ("select ID,file from Candidate where Name='" + Combo1(Index).Text + "'")
    If IsNull(t.rs.Fields(1)) Or t.rs.Fields(1) = "" Then
    Image1(Index).Picture = LoadPicture(App.Path & "\vo_img\base.gif")
    Else
    Image1(Index).Picture = LoadPicture(App.Path & "\vo_img\" & t.rs.Fields(1))
    End If
    Text1(Index).Text = t.rs.Fields(0)
End If
t.rs.Close
t.db.Close
End Sub

Private Sub Command1_Click(Index As Integer)
Select Case Index
    Case 0
        frmvote_party.Show vbModal, Me
    Case 1
        For i = 0 To 7
            Combo1(i).Text = ""
        Next
    Case 2
        frmconfirm.Show vbModal, Me
End Select
End Sub

Private Sub Form_Load()
t.Database ("Select * from Post")
With t.rs
    .MoveFirst
    While Not t.rs.EOF
        Call hide1(True, .Fields(0) - 1)
        Label3(.Fields(0) - 1).Caption = .Fields(1)
        t.rs.MoveNext
    Wend
End With
t3.Database ("select Class,Year from Voter Where V_ID='" + Login.Text5.Text + "'")
t.rs.MoveFirst
While Not t.rs.EOF
If t.rs.Fields("Level") = "College" Then
    t2.Database ("select Name from Candidate where Post='" + t.rs.Fields(1) + "'")
        While Not t2.rs.EOF
        Combo1(t.rs.Fields(0) - 1).AddItem t2.rs.Fields(0)
        t2.rs.MoveNext
        Wend
    t2.rs.Close
    t2.db.Close
ElseIf t.rs.Fields("Level") = "Class" Then
    t2.Database ("select Name from Candidate where Post='" + t.rs.Fields(1) + "' and Class='" + t3.rs.Fields(0) + "'and Year ='" + t3.rs.Fields(1) + "'")
        While Not t2.rs.EOF
            Combo1(t.rs.Fields(0) - 1).AddItem t2.rs.Fields(0)
            t2.rs.MoveNext
        Wend
    t2.rs.Close
    t2.db.Close
End If
    t.rs.MoveNext
Wend
t.rs.Close
t.db.Close
t3.rs.Close
t3.db.Close
Label4.Caption = "Welcome " & Login.Text5.Text
Image2.Width = 80
Image2.Height = 80
Image2.Top = 1440
Image2.Left = 1440
End Sub

Private Sub Form_Unload(Cancel As Integer)
m = MsgBox("Are you sure?", vbQuestion + vbApplicationModal + vbYesNo)
If m = vbYes Then
Unload Me
Login.Show
Login.Text5.Text = ""
Login.Text7.Text = ""
Login.Combo4.Text = ""
Login.Text5.SetFocus
Else
Cancel = 1
End If
End Sub

Private Sub Timer1_Timer()
Image2.Width = Image2.Width + 100
Image2.Height = Image2.Height + 100
Image2.Top = Image2.Top - 50
Image2.Left = Image2.Left - 50
If Image2.Width = 2880 Then
Timer1.Enabled = False
Image2.Left = -10
Image2.Top = 240
End If
End Sub
