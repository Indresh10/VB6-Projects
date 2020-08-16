VERSION 5.00
Begin VB.Form frmconfirm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Confirm your Votes"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11700
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cancel"
      Height          =   615
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Vote"
      Height          =   615
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   7
      Left            =   10440
      TabIndex        =   16
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   6
      Left            =   10440
      TabIndex        =   15
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   5
      Left            =   10440
      TabIndex        =   14
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   4
      Left            =   10440
      TabIndex        =   13
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   3
      Left            =   4800
      TabIndex        =   12
      Top             =   3360
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   480
      Index           =   2
      Left            =   4800
      TabIndex        =   11
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label3"
      ForeColor       =   &H80000008&
      Height          =   495
      Index           =   0
      Left            =   4800
      TabIndex        =   9
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   7
      Left            =   6120
      TabIndex        =   8
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   6
      Left            =   6120
      TabIndex        =   7
      Top             =   2475
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   5
      Left            =   6120
      TabIndex        =   6
      Top             =   1605
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   4
      Left            =   6120
      TabIndex        =   5
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3360
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   2475
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   1605
      Width           =   840
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "lghgh"
      Height          =   540
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Your Chosen Candidates are"
      Height          =   540
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4485
   End
End
Attribute VB_Name = "frmconfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Dim rs As ADODB.Recordset
Private Sub Command1_Click()
Screen.MousePointer = 11
Set t.db = New ADODB.Connection
Source = App.Path & "\all.mdb"
t.db.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Source
t.db.Open
t.db.Execute ("ALTER TABLE Votes ADD " + Trim(Login.Text5.Text) + " INTEGER")
t.db.Close
Set rs = New ADODB.Recordset
t.Database ("Select * from Votes")
t.rs.MoveFirst
For i = 0 To 7
    If frmvote.Label3(i).Visible = True Then
        With t.rs
            .Fields(Trim(Login.Text5.Text)) = Val(frmvote.Text1(i).Text)
            .MoveNext
        End With
        rs.Open "Select Votes from Candidate where ID=" & Val(frmvote.Text1(i).Text), t.db, adOpenKeyset, adLockOptimistic
            If Not rs.EOF Then: rs.Fields(0) = rs.Fields(0) + 1: rs.Update
        rs.Close
    End If
Next
If t.rs.EOF Then: t.rs.MoveLast: t.rs.Update
t.rs.Close
rs.Open "select Voted From Voter where V_ID='" + Trim(Login.Text5.Text) + "'", t.db, adOpenKeyset, adLockOptimistic
    If Not rs.EOF Then: rs.Fields(0) = "Yes": rs.Update
rs.Close
t.db.Close
Unload Me
Unload frmvote
Screen.MousePointer = 0
MsgBox "Your vote has been added to the system", vbInformation
t.Database ("Select C.Name,Votes.Post from Votes LEFT JOIN Candidate as C ON C.ID=Votes." & Trim(Login.Text5.Text))
Set votereport.DataSource = t.rs
votereport.Show
Login.Show
Login.Text5.Text = ""
Login.Text7.Text = ""
Login.Combo4.Text = ""
Login.Text5.SetFocus
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
For i = 0 To 7
    If frmvote.Label3(i).Visible = True Then
        Label2(i).Visible = True
        Label3(i).Visible = True
        Label2(i).Caption = frmvote.Label3(i).Caption
        Label3(i).Caption = frmvote.Combo1(i).Text
    Else
        Label2(i).Visible = False
        Label3(i).Visible = False
    End If
Next
End Sub

