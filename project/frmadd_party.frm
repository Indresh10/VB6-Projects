VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmadd_party 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Party"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7860
   BeginProperty Font 
      Name            =   "Kristen ITC"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmadd_party.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4305
      ScaleWidth      =   7545
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3960
         TabIndex        =   6
         Top             =   1320
         Width           =   2895
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Upload  Photo"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   3000
         Width           =   2295
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   3000
         Width           =   2295
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   2175
         Left            =   120
         Stretch         =   -1  'True
         Top             =   720
         Width           =   2295
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Add Party"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   7215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   1440
         Width           =   975
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   6480
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmadd_party"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Dim file As String
Private Sub Command1_Click()
On Error GoTo cdhandler
start:
With CD
    .Filter = "JPEG Files(*.jpg)|*.jpg|GIF Files(*.gif)|*.gif"
    .ShowOpen
End With
fn = "test.jpg"
If Not CD.FileName = "" Then
FileCopy CD.FileName, App.Path & "\party_logo\" + fn
file = App.Path & "\party_logo\" + fn
Image1.Picture = LoadPicture(file)
End If
Exit Sub
cdhandler:
Select Case MsgBox(Error(Err.Number), vbCritical + vbAbortRetryIgnore, "Error number-" & Str(Err.Number))
Case vbAbort
    Resume exitline
Case vbRetry
    Resume start
Case vbIgnore
    Resume Next
End Select
exitline:
End Sub


Private Sub Command2_Click()
If Image1.Picture = none Then
MsgBox "Your photo Is not uploaded", vbExclamation
Exit Sub
End If
If Text1.Text = "" Then
MsgBox "Please Enter the Party name", vbExclamation
Exit Sub
End If
t.Database ("Select * from Party")
If t.rs.BOF And t.rs.EOF Then
    id = 1
Else
    t.rs.MoveLast
    id = t.rs.Fields(0) + 1
End If
f = Text1.Text
fn = f & id & ".jpg"
Name file As App.Path + "\party_logo\" + fn
With t.rs
    .AddNew
    .Fields(0) = id
    .Fields(1) = Text1.Text
    .Fields(2) = fn
    .Update
    .Close
End With
file = ""
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
If file <> "" Then Kill file
End Sub
