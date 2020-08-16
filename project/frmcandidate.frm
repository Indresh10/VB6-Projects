VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmcand_man 
   BackColor       =   &H00FF8080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Candidate Management"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   12975
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmcandidate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   12975
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      BackColor       =   &H008080FF&
      Caption         =   "Descending"
      Height          =   480
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5450
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H008080FF&
      Caption         =   "Ascending"
      Height          =   480
      Left            =   4440
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   5450
      Width           =   1935
   End
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFF80&
      Height          =   975
      Left            =   11880
      Picture         =   "frmcandidate.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5040
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   3945
      TabIndex        =   18
      Top             =   300
      Width           =   3975
      Begin VB.CommandButton Command1 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ADD New Party"
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "DELETE a Party"
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command4 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   " CLEAR Party List"
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   3000
         Width           =   2655
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Party"
         Height          =   495
         Left            =   120
         TabIndex        =   22
         Top             =   120
         Width           =   3735
      End
      Begin VB.Image Image1 
         Height          =   900
         Left            =   240
         Picture         =   "frmcandidate.frx":1820
         Top             =   840
         Width           =   900
      End
      Begin VB.Image Image2 
         Height          =   900
         Left            =   120
         Picture         =   "frmcandidate.frx":1F76
         Top             =   1920
         Width           =   900
      End
      Begin VB.Image Image3 
         Height          =   900
         Left            =   240
         Picture         =   "frmcandidate.frx":25DD
         Top             =   3000
         Width           =   900
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3375
      Left            =   4200
      TabIndex        =   0
      Top             =   1560
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5953
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      DefColWidth     =   133
      HeadLines       =   1
      RowHeight       =   26
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   4200
      ScaleHeight     =   3105
      ScaleWidth      =   8625
      TabIndex        =   7
      Top             =   120
      Width           =   8655
      Begin VB.ComboBox Combo1 
         Height          =   480
         Left            =   2280
         TabIndex        =   10
         Top             =   840
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   495
         Left            =   5040
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         Height          =   480
         Left            =   5040
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Image Image11 
         Height          =   615
         Left            =   7920
         Picture         =   "frmcandidate.frx":2B9F
         Stretch         =   -1  'True
         ToolTipText     =   "refresh Database"
         Top             =   720
         Width           =   615
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   7440
         Picture         =   "frmcandidate.frx":2CB8
         Stretch         =   -1  'True
         Top             =   840
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Search"
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   8415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH BY "
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   840
         Width           =   2055
      End
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   4200
      ScaleHeight     =   2865
      ScaleWidth      =   8625
      TabIndex        =   4
      Top             =   6120
      Width           =   8655
      Begin VB.CommandButton Command9 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "CLEAR Candidate List"
         Height          =   975
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton Command8 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "DELETE a Candidate"
         Height          =   975
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   1680
         Width           =   2655
      End
      Begin VB.CommandButton Command5 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ADD New Candidate "
         Height          =   975
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Image Image10 
         Height          =   900
         Left            =   6840
         Picture         =   "frmcandidate.frx":31ED
         Top             =   720
         Width           =   900
      End
      Begin VB.Image Image9 
         Height          =   900
         Left            =   3840
         Picture         =   "frmcandidate.frx":37AF
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Candidate"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   8415
      End
      Begin VB.Image Image5 
         Height          =   900
         Left            =   960
         Picture         =   "frmcandidate.frx":3E16
         Top             =   720
         Width           =   900
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4185
      ScaleWidth      =   3945
      TabIndex        =   1
      Top             =   4620
      Width           =   3975
      Begin VB.CommandButton Command7 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   " CLEAR Post List"
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   3000
         Width           =   2655
      End
      Begin VB.CommandButton Command6 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "DELETE a Post"
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1920
         Width           =   2655
      End
      Begin VB.CommandButton Command3 
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "ADD New Post "
         Height          =   975
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.Image Image8 
         Height          =   900
         Left            =   240
         Picture         =   "frmcandidate.frx":4533
         Top             =   3000
         Width           =   900
      End
      Begin VB.Image Image7 
         Height          =   900
         Left            =   120
         Picture         =   "frmcandidate.frx":4AF5
         Top             =   1920
         Width           =   900
      End
      Begin VB.Image Image6 
         Height          =   900
         Left            =   240
         Picture         =   "frmcandidate.frx":515C
         Top             =   840
         Width           =   900
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Post "
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFC0&
      BackStyle       =   0  'Transparent
      Caption         =   "Print Candidate -->"
      Height          =   360
      Left            =   8760
      TabIndex        =   24
      Top             =   5400
      Width           =   2955
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Sort By ID:"
      Height          =   975
      Left            =   4200
      TabIndex        =   17
      Top             =   5040
      Width           =   8655
   End
End
Attribute VB_Name = "frmcand_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Private Function datagrid()
Set DataGrid1.DataSource = t.rs
DataGrid1.Columns(0).Width = 750
DataGrid1.Columns(1).Width = 1750
DataGrid1.Columns(3).Visible = False
DataGrid1.Columns(4).Visible = False
DataGrid1.Columns(5).Width = 1750
End Function
Private Function Search()
Select Case Combo1.Text
Case "Name"
    Combo2.Visible = False
    Text1.Visible = True
    Image4.Visible = True
Case "Post"
    t.Database ("Select * from Post order by ID ASC")
    Combo2.Visible = True
    Text1.Visible = False
    Image4.Visible = True
    Combo2.clear
    While Not t.rs.EOF
        Combo2.AddItem t.rs.Fields(1)
        t.rs.MoveNext
    Wend
    t.rs.Close
Case "Party"
    t.Database ("Select * From Party Order By ID ASC")
    Combo2.Visible = True
    Text1.Visible = False
    Image4.Visible = True
    Combo2.clear
    While Not t.rs.EOF
        Combo2.AddItem t.rs.Fields(1)
        t.rs.MoveNext
    Wend
    t.rs.Close
Case Else
    Text1.Visible = False
    Combo2.Visible = False
    Image4.Visible = False
    t.Database ("Select * from sub_candidate")
    m = datagrid()
End Select
End Function
Private Sub Combo1_Change()
m = Search()
End Sub
Private Sub Combo1_Click()
m = Search()
End Sub
Private Sub Combo2_GotFocus()
Image4.Picture = LoadPicture(App.Path & "\logo\search2.gif")
End Sub

Private Sub Combo2_LostFocus()
Image4.Picture = LoadPicture(App.Path & "\logo\search.gif")
End Sub

Private Sub Command1_Click()
frmadd_party.Show vbModal, Me
Combo1.Text = ""
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\logo\add.gif")
End Sub


Private Sub Command10_Click()
Dim year As String
Dim month As Integer
t.Database ("Select * from Candidate")
Set canreport.DataSource = DataGrid1.DataSource
month = Format(Date, "MM")
year = Format(Date, "YYYY")
If month > 6 Then
year = year + "-" & (Val(Right(year, 2)) + 1)
Else
year = Val(year) - 1 & "-" + Right(year, 2)
End If
canreport.Sections(2).Controls("Label2").Caption = year
canreport.Show vbModal, Me
End Sub

Private Sub Command2_Click()
frmdelete.Text1.Text = "Party"
frmdelete.Show vbModal, Me
t.Database ("Select * from sub_candidate")
t.rs.Requery
n = datagrid()
End Sub

Private Sub Command2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Picture = LoadPicture(App.Path & "\logo\edit2.gif")
End Sub

Private Sub Command3_Click()
frmadd_post.Show vbModal, Me
End Sub

Private Sub Command3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Picture = LoadPicture(App.Path & "\logo\add.gif")
End Sub

Private Sub Command4_Click()
out = MsgBox("Are You Sure ?", vbYesNo + vbExclamation)
If out = 6 Then
t.db.Execute ("Delete from Party where ID not like '1'")
t.rs.Requery
Combo1.Text = ""
End If
End Sub

Private Sub Command4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3.Picture = LoadPicture(App.Path & "\logo\delete2.gif")
End Sub

Private Sub Command5_Click()
frmadd_cand.Show vbModal, Me
t.rs.Requery
t.Database ("Select * from sub_candidate")
t.rs.Requery
n = datagrid()
End Sub

Private Sub Command5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\logo\add_candidate.gif")
End Sub

Private Sub Command6_Click()
frmdelete.Text1.Text = "Post"
frmdelete.Show vbModal, Me
t.Database ("Select * from sub_candidate")
t.rs.Requery
n = datagrid()
End Sub

Private Sub Command6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image7.Picture = LoadPicture(App.Path & "\logo\edit2.gif")
End Sub

Private Sub Command7_Click()
out = MsgBox("Are You Sure ?", vbYesNo + vbExclamation)
If out = 6 Then
t.db.Execute ("Delete from Post")
t.db.Execute ("Delete from Votes")
Combo1.Text = ""
End If
End Sub

Private Sub Command7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image8.Picture = LoadPicture(App.Path & "\logo\delete2.gif")
End Sub

Private Sub Command8_Click()
m = Val(InputBox("Enter ID of candidate"))
t.Database ("select * from candidate where ID=" & m)
If Not t.rs.EOF Then
fname = App.Path + "\vo_img\" + t.rs.Fields(3)
Kill fname
t.db.Execute ("Delete from Candidate where ID=" & m)
t.Database ("Select * from sub_candidate")
t.rs.Requery
n = datagrid()
End If
End Sub

Private Sub Command8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image9.Picture = LoadPicture(App.Path & "\logo\edit2.gif")
End Sub

Private Sub Command9_Click()
m = MsgBox("Are You Sure ?", vbYesNo + vbExclamation)
t.Database ("select * from candidate")
If Not t.rs.EOF Then
While Not t.rs.EOF
fname = App.Path + "\vo_img\" + t.rs.Fields("file")
If Not IsNull(fname) Then Kill fname
t.rs.MoveNext
Wend
End If
t.rs.Close
t.Database ("select * from sub_candidate")
If m = 6 Then
t.db.Execute ("Delete from Candidate")
t.rs.Requery
n = datagrid()
End If
End Sub

Private Sub Command9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image10.Picture = LoadPicture(App.Path & "\logo\delete2.gif")
End Sub

Public Sub Form_Load()
Combo1.AddItem "Name"
Combo1.AddItem "Party"
Combo1.AddItem "Post"
t.Database ("Select * from sub_candidate")
m = datagrid()
End Sub

Private Sub Image11_Click()
t.Database ("Select * from sub_candidate")
t.rs.Requery
n = datagrid()
Combo1.Text = ""
Option1.Value = False
Option2.Value = False
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Picture = LoadPicture(App.Path & "\logo\datarefresh2.gif")
End Sub

Private Sub Image4_Click()
Select Case Combo1.Text
Case "Name"
    t.Database ("Select * From sub_candidate where Name='" + Text1.Text + "'")
    
Case "Post"
    t.Database ("Select * From sub_candidate where Post='" + Combo2.Text + "'")
  
Case "Party"
    t.Database ("Select * From sub_candidate where Party='" + Combo2.Text + "'")
    
Case Else
    MsgBox "Select Category first", vbExclamation
End Select
m = datagrid()
End Sub

Private Sub Option1_Click()
t.Database ("Select * from sub_candidate order by ID")
m = datagrid()
End Sub

Private Sub Option2_Click()
t.Database ("Select * from sub_candidate order by ID Desc")
m = datagrid()
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\logo\add2.gif")
Image2.Picture = LoadPicture(App.Path & "\logo\edit.gif")
Image3.Picture = LoadPicture(App.Path & "\logo\delete.gif")
End Sub
Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image6.Picture = LoadPicture(App.Path & "\logo\add2.gif")
Image7.Picture = LoadPicture(App.Path & "\logo\edit.gif")
Image8.Picture = LoadPicture(App.Path & "\logo\delete.gif")
End Sub
Private Sub Picture3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image5.Picture = LoadPicture(App.Path & "\logo\add_candidate2.gif")
Image9.Picture = LoadPicture(App.Path & "\logo\edit.gif")
Image10.Picture = LoadPicture(App.Path & "\logo\delete.gif")
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Picture = LoadPicture(App.Path & "\logo\datarefresh.gif")
End Sub

Private Sub Text1_GotFocus()
Image4.Picture = LoadPicture(App.Path & "\logo\search2.gif")
End Sub

Private Sub Text1_LostFocus()
Image4.Picture = LoadPicture(App.Path & "\logo\search.gif")
End Sub
