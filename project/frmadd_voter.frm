VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmadd_voter 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Voter Registration"
   ClientHeight    =   6105
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4770
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Ink Free"
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
   ScaleHeight     =   6105
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      Caption         =   "Voter Registration"
      Height          =   6015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox Text1 
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
         Left            =   2040
         TabIndex        =   13
         Top             =   4200
         Width           =   2535
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   2535
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cancel"
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
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5280
         Width           =   1455
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Submit"
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
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5280
         Width           =   1455
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2040
         TabIndex        =   2
         Top             =   2655
         Width           =   2535
      End
      Begin VB.TextBox Text6 
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
         Left            =   2040
         TabIndex        =   1
         Top             =   960
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   2040
         TabIndex        =   11
         Top             =   3480
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   133431297
         CurrentDate     =   43817
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Adm. No"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   12
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Voter Registration"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   2925
      End
      Begin VB.Image Image1 
         Height          =   720
         Left            =   3960
         Picture         =   "frmadd_voter.frx":0000
         Top             =   0
         Width           =   720
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOB"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   9
         Top             =   3600
         Width           =   735
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   8
         Top             =   2760
         Width           =   720
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   7
         Top             =   1920
         Width           =   765
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   240
         TabIndex        =   6
         Top             =   1080
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmadd_voter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Dim rs1 As New ADODB.Recordset
Private Function v_id(ByVal name, ByVal Class, ByVal year, ByVal dob, ByVal id) As String
Dim vid As String
vid = "CC10"
vid = vid + Left$(name, 2)
t.Database ("select ID from Courses where Name='" + Class + "'")
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

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Dim name, dt As String
Dim id As Integer
If name = "" Or Text1.Text = "" Or Combo3.Text = "" Or Combo2.Text = "" Then
MsgBox "Please Fill All the details", vbExclamation
Exit Sub
End If
name = Text6.Text
dt = Format(DTPicker1.Value, "dd-mm-yyyy")
If name = "" Or Combo3.Text = "" Or Combo2.Text = "" Or Text1.Text = "" Then
    MsgBox "please complete your details", vbExclamation
    Exit Sub
End If
t.Database ("SELECT * FROM Voter WHERE Name='" + Trim(name) + "' and DOB=#" + dt + "#")
rs1.Open ("SELECT * FROM Voter"), t.db, adOpenKeyset, adLockOptimistic
If rs1.BOF And rs1.EOF Then
    id = 1
Else
    rs1.MoveLast
    id = rs1.Fields(0) + 1
End If
If t.rs.EOF Then
    vid = v_id(Text6.Text, Combo3.Text, Combo2.Text, dt, id)
    pass = Left(name, 2) + Right(Trim(Text1.Text), 4)
    With rs1
        .AddNew
        .Fields(0) = id
        .Fields(1) = UCase$(vid)
        .Fields(2) = name
        .Fields(3) = Combo3.Text
        .Fields(4) = Combo2.Text
        .Fields(5) = dt
        .Fields(6) = pass
        .Fields(8) = "No"
        .Fields(9) = pass
        .Update
    End With
    MsgBox "Record Added Sucessfully", vbInformation, "Registration"
    MsgBox "Voter id is " + UCase$(vid) + vbCr + "Password is " + pass, vbInformation, "Voter ID"
    Call clear
    rs1.Requery
Else
    MsgBox "Already Registered", vbExclamation, "Registration"
    Call clear
End If
rs1.Close
End Sub

Private Sub Form_Load()
Combo2.AddItem "First"
Combo2.AddItem "Second"
Combo2.AddItem "Third"
t.Database ("Select Name from Courses")
While Not t.rs.EOF
    Combo3.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
t.rs.Close
t.db.Close
DTPicker1.Value = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\logo\close.gif")
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\logo\close2.gif")
End Sub

Private Sub Text4_Validate(Cancel As Boolean)
If Text4.Text <> Text3.Text Then
    MsgBox "Please Check Password", vbExclamation, "Registration"
    Text3.Text = ""
    Text4.Text = ""
    Text3.SetFocus
End If
End Sub

Private Function clear()
Text6.Text = ""
Combo3.Text = ""
Text1.Text = ""
Combo2.Text = ""
DTPicker1.Value = Format(Now, "dd-mm-yyyy")
Text6.SetFocus
End Function
