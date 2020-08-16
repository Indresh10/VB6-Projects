VERSION 5.00
Begin VB.Form frmdelete 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   5385
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   2400
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   840
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   1680
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancel"
         Height          =   1095
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete"
         Height          =   1095
         Left            =   360
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Delete"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmdelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Private Sub Command2_Click()
m = Combo1.Text
If Not m = "" Then
If Text1.Text = "Party" Then
    t.Database ("Select * From Party Order by ID")
    t.rs.Find ("Name='" + m + "'")
    If Not (t.rs.EOF Or m = "Independent") Then
    t.db.Execute ("delete from Candidate where Party='" + m + "'")
    n = t.rs.Fields(0)
    t.rs.Delete
    t.rs.MoveFirst
    t.rs.Find ("ID=" & n + 1)
    While Not t.rs.EOF
    t.rs.Update 0, t.rs.Fields(0) - 1
    t.rs.MoveNext
    Wend
    MsgBox "Succesfully deleted " & m, vbInformation
    t.rs.Close
    t.db.Close
    Unload Me
    End If
Else
    t.Database ("Select * From Post order by ID")
    t.rs.Find ("Name='" + m + "'")
    If Not t.rs.EOF Then
    t.db.Execute ("delete from Votes where Post='" + m + "'")
    t.db.Execute ("delete from Candidate where Post='" + m + "'")
    n = t.rs.Fields(0)
    t.rs.Delete
    t.rs.MoveFirst
    t.rs.Find ("ID=" & n + 1)
    While Not t.rs.EOF
    t.rs.Update 0, t.rs.Fields(0) - 1
    t.rs.MoveNext
    Wend
    MsgBox "Succesfully deleted " & m, vbInformation
    t.rs.Close
    t.db.Close
    End If
    Unload Me
End If
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Text1_Change()
t.Database ("Select * From " + Text1.Text)
Combo1.clear
While Not t.rs.EOF
    Combo1.AddItem t.rs.Fields(1)
    t.rs.MoveNext
Wend
End Sub
