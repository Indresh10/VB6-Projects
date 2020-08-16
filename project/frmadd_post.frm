VERSION 5.00
Begin VB.Form frmadd_post 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Post"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadd_post.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   120
      ScaleHeight     =   4065
      ScaleWidth      =   5265
      TabIndex        =   0
      Top             =   120
      Width           =   5295
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1920
         Width           =   2895
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2880
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1920
         TabIndex        =   1
         Top             =   960
         Width           =   2895
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Level"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   6
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Add Post"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   5055
      End
   End
End
Attribute VB_Name = "frmadd_post"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1

Private Sub Command2_Click()
If Text1.Text = "" Or Combo1.Text = "" Then
MsgBox "Please Fill the Details", vbExclamation
Exit Sub
End If
t.Database ("Select * from Post")
If t.rs.BOF And t.rs.EOF Then
    id = 1
Else
    t.rs.MoveLast
    id = t.rs.Fields(0) + 1
End If
If id <= 8 Then
With t.rs
    .AddNew
    .Fields(0) = id
    .Fields(1) = Text1.Text
    .Fields(2) = Combo1.Text
    .Update
    .Close
End With
t.db.Close
t.Database ("Select * from Votes")
If t.rs.BOF And t.rs.EOF Then
    id = 1
Else
    t.rs.MoveLast
    id = t.rs.Fields(0) + 1
End If
With t.rs
    .AddNew
    .Fields(0) = id
    .Fields(1) = Text1.Text
    .Update
    .Close
End With
t.db.Close
MsgBox "Succesfully Added Post"
Unload Me
Else
MsgBox "Sorry Post Limit is only 8", vbInformation
End If
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1.AddItem "College"
Combo1.AddItem "Class"
End Sub
