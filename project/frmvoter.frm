VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmvoter 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Voter Maintenance "
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13125
   BeginProperty Font 
      Name            =   "Lucida Handwriting"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmvoter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8865
   ScaleWidth      =   13125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command10 
      BackColor       =   &H00FFFF80&
      Height          =   975
      Left            =   7920
      Picture         =   "frmvoter.frx":10CA
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7800
      Width           =   975
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   4080
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   5880
      Top             =   3480
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   9000
      ScaleHeight     =   5625
      ScaleWidth      =   3945
      TabIndex        =   9
      Top             =   3120
      Width           =   3975
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password Reset to Default"
         Height          =   1215
         Index           =   3
         Left            =   1440
         TabIndex        =   14
         Top             =   4200
         Width           =   1860
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   600
         Index           =   3
         Left            =   480
         Picture         =   "frmvoter.frx":1820
         Top             =   4560
         Width           =   600
      End
      Begin VB.Image Image2 
         Height          =   900
         Index           =   2
         Left            =   360
         Picture         =   "frmvoter.frx":1CAE
         Top             =   3000
         Width           =   900
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Register Student"
         Height          =   810
         Index           =   2
         Left            =   1440
         TabIndex        =   13
         Top             =   3120
         Width           =   1515
         WordWrap        =   -1  'True
      End
      Begin VB.Image Image2 
         Height          =   1500
         Index           =   1
         Left            =   0
         Picture         =   "frmvoter.frx":23CB
         Top             =   1500
         Width           =   1500
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reset Status"
         Height          =   405
         Index           =   1
         Left            =   1440
         TabIndex        =   12
         Top             =   2040
         Width           =   2070
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delete"
         Height          =   405
         Index           =   0
         Left            =   1440
         TabIndex        =   11
         Top             =   960
         Width           =   1200
      End
      Begin VB.Image Image2 
         Height          =   900
         Index           =   0
         Left            =   240
         Picture         =   "frmvoter.frx":2AE2
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Options"
         Height          =   495
         Left            =   120
         TabIndex        =   10
         Top             =   120
         Width           =   3735
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   9000
      ScaleHeight     =   2865
      ScaleWidth      =   3945
      TabIndex        =   5
      Top             =   120
      Width           =   3975
      Begin VB.ComboBox Combo2 
         Height          =   525
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   2055
      End
      Begin VB.ComboBox Combo1 
         Height          =   525
         Left            =   480
         TabIndex        =   8
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   1680
         Picture         =   "frmvoter.frx":30A4
         Stretch         =   -1  'True
         ToolTipText     =   "refresh Database"
         Top             =   2160
         Width           =   615
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Filter by"
         Height          =   405
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1560
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Filter"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   3735
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7215
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   12726
      _Version        =   393216
      AllowUpdate     =   0   'False
      BackColor       =   16777215
      DefColWidth     =   167
      HeadLines       =   1
      RowHeight       =   26
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Script"
         Size            =   9.75
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
      Left            =   120
      ScaleHeight     =   3105
      ScaleWidth      =   8745
      TabIndex        =   0
      Top             =   120
      Width           =   8775
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   3360
         TabIndex        =   3
         Top             =   840
         Width           =   3975
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         Visible         =   0   'False
         X1              =   3360
         X2              =   3375
         Y1              =   1320
         Y2              =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         BorderWidth     =   3
         X1              =   3360
         X2              =   7320
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SEARCH BY NAME"
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   2985
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H0080FF80&
         Caption         =   "Search"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   8535
      End
      Begin VB.Image Image4 
         Height          =   495
         Left            =   7440
         Picture         =   "frmvoter.frx":31BD
         Stretch         =   -1  'True
         Top             =   840
         Width           =   495
      End
      Begin VB.Image Image11 
         Height          =   615
         Left            =   8040
         Picture         =   "frmvoter.frx":36F2
         Stretch         =   -1  'True
         ToolTipText     =   "refresh Database"
         Top             =   720
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmvoter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Public Function datagrid(ByVal db As String)
t.Database (db)
Set DataGrid1.DataSource = t.rs
DataGrid1.Refresh
DataGrid1.Columns(1).Width = 2150
DataGrid1.Columns(2).Width = 750
DataGrid1.Columns(3).Width = 800
DataGrid1.Columns(4).Width = 1500
DataGrid1.Columns(5).Width = 700
End Function


Private Sub Combo1_Change()
If Not Combo1.Text = "" Then: datagrid ("select * from sub_voter where " + Combo2.Text + "='" + Combo1.Text + "'")
End Sub

Private Sub Combo1_Click()
datagrid ("select * from sub_voter where " + Combo2.Text + "='" + Combo1.Text + "'")
Picture1.SetFocus
End Sub




Private Sub Combo2_Click()
Combo1.clear
t.Database ("select distinct " + Combo2.Text + " from Voter")
While Not t.rs.EOF
    Combo1.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
End Sub

Private Sub Command10_Click()
Set voterreport.DataSource = DataGrid1.DataSource
voterreport.Show vbModal, Me
End Sub

Private Sub Form_Load()
t.Database ("Select Name from Courses")
While Not t.rs.EOF
    Combo1.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
t.rs.Close
Combo2.AddItem "Class"
Combo2.AddItem "Year"
Combo2.AddItem "Voted"
datagrid ("Select * From sub_voter")
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub

Private Sub Image1_Click()
datagrid ("Select * from sub_voter")
Text1.Text = ""
Combo1.Text = ""
Picture1.SetFocus
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\logo\datarefresh2.gif")
End Sub

Private Sub Image11_Click()
datagrid ("Select * from sub_voter")
Text1.Text = ""
Combo1.Text = ""
Picture4.SetFocus
End Sub

Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image11.Picture = LoadPicture(App.Path & "\logo\datarefresh2.gif")
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\logo\search2.gif")
End Sub

Private Sub Image4_Click()
m = datagrid("Select * from sub_voter where Name='" + Text1.Text + "'")
Picture4.SetFocus
End Sub




Private Sub Label6_Click(Index As Integer)
Select Case Index
    Case 0 'delete
        m = InputBox("Enter Voter ID")
        t.Database ("Select * from Voter where V_ID='" + m + "'")
        If Not t.rs.EOF Then
        t.db.Execute ("delete from Voter where V_ID='" + m + "'")
        t.rs.Requery
        MsgBox "Succesfully deleted " + m, vbInformation
        datagrid ("select * from sub_voter")
        End If
    Case 1
        m = InputBox("Enter VoterID")
        t.Database ("select * from Voter where V_ID='" + m + "'")
        If Not t.rs.EOF Then
        If t.rs.Fields(7) = "Yes" Then
            t.rs.Fields(7) = "No"
            t.rs.Update
            MsgBox "Status changed succesfully", vbInformation
        End If
        t.rs.Close
        t.db.Execute ("ALTER TABLE Votes DROP COLUMN " + m)
        Call Image11_Click
        End If
    Case 2
        frmadd_voter.Show vbModal, Me
    Case 3
        m = InputBox("Enter Voter ID", "Default")
        t.Database ("Select * from Voter where V_ID='" + m + "'")
        If t.rs.EOF Then
            MsgBox "Please enter a valid Voter ID", vbExclamation
        Else
            t.rs.Update "Password", t.rs.Fields("Default")
            MsgBox "password changed back to default", vbInformation
        End If
    End Select
End Sub

Private Sub Label6_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        Image2(Index).Picture = LoadPicture(App.Path & "\logo\delete2.gif")
    Case 1
        Image2(Index).Picture = LoadPicture(App.Path & "\logo\refresh2.gif")
    Case 2
        Image2(Index).Picture = LoadPicture(App.Path & "\logo\add_candidate.gif")
    Case 3
        Image2(Index).Picture = LoadPicture(App.Path & "\logo\show_pass2.gif")
End Select
Label6(Index).ForeColor = vbRed
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Picture = LoadPicture(App.Path & "\logo\datarefresh.gif")
End Sub


Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(0).Picture = LoadPicture(App.Path & "\logo\delete.gif")
Image2(1).Picture = LoadPicture(App.Path & "\logo\refresh.gif")
Image2(2).Picture = LoadPicture(App.Path & "\logo\add_candidate2.gif")
Image2(3).Picture = LoadPicture(App.Path & "\logo\show_pass.gif")
For i = 0 To 3
    Label6(i).ForeColor = vbBlack
Next
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image4.Picture = LoadPicture(App.Path & "\logo\search.gif")
Image11.Picture = LoadPicture(App.Path & "\logo\datarefresh.gif")
End Sub
Private Sub Text1_GotFocus()
Line2.Visible = True
Timer1.Enabled = True
Line2.X1 = 3360
Line2.X2 = 3360
Line2.Y1 = 1320
Line2.Y2 = 1320
Timer2.Enabled = False
End Sub

Private Sub Text1_LostFocus()
Timer2.Enabled = True
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
Line2.X2 = Line2.X2 + 60
If Line2.X2 = 7320 Then: Timer1.Enabled = False
End Sub

Private Sub Timer2_Timer()
Line2.X2 = Line2.X2 - 60
If Line2.X2 = 3360 Then
    Line2.Visible = False
    Timer2.Enabled = False
End If
End Sub
