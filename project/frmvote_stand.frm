VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmvote_stand 
   BackColor       =   &H0060FF60&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Votes Standing"
   ClientHeight    =   8175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   14790
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmvote_stand.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8175
   ScaleWidth      =   14790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   12000
      ScaleHeight     =   2145
      ScaleWidth      =   2625
      TabIndex        =   6
      Top             =   2760
      Width           =   2655
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         TabIndex        =   9
         Top             =   1320
         Width           =   1815
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00AFFFF0&
         Caption         =   "Filter By"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   12000
      ScaleHeight     =   2505
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton Option2 
         BackColor       =   &H00A0A0FF&
         Caption         =   "Descending"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00A0A0FF&
         Caption         =   "Ascending"
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   480
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Lucida Calligraphy"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00AFFFF0&
         Caption         =   "Sort By"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Options"
      ForeColor       =   &H80000008&
      Height          =   3015
      Left            =   12000
      TabIndex        =   10
      Top             =   5040
      Width           =   2655
      Begin VB.CommandButton Command3 
         BackColor       =   &H00BBFF00&
         Caption         =   "Stop Election"
         Height          =   855
         Left            =   480
         MaskColor       =   &H00C0C0FF&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   480
         UseMaskColor    =   -1  'True
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00BBFF00&
         Caption         =   "Winners"
         Height          =   525
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   2280
         Width           =   1695
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00BBFF00&
         Caption         =   "Refresh"
         Height          =   525
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   1695
      End
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00BBFF00&
      Caption         =   "Start Election"
      Height          =   1095
      Left            =   5888
      MaskColor       =   &H00C0C0FF&
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2865
      UseMaskColor    =   -1  'True
      Width           =   3015
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00BBFF00&
      Caption         =   "Last Election Winners"
      Height          =   1005
      Left            =   5888
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4305
      Width           =   3015
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   7935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   13996
      _Version        =   393216
      AllowUpdate     =   -1  'True
      DefColWidth     =   117
      HeadLines       =   1
      RowHeight       =   29
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Print"
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
End
Attribute VB_Name = "frmvote_stand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Dim t2 As New Database1
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim year(3) As String
Private Function datagrid()
Set DataGrid1.DataSource = t.rs
DataGrid1.Columns(5).Visible = False
DataGrid1.Columns(0).Width = 800
DataGrid1.Columns(2).Width = 1850
DataGrid1.Columns(4).Width = 850
DataGrid1.Columns(8).Width = 800
End Function
Private Sub Combo1_Click()
Option1.Value = False
Option2.Value = False
End Sub

Private Sub Combo2_Click()
Combo3.Text = ""
Combo3.clear
t.Database ("SELECT DISTINCT " & Combo2.Text & " FROM Candidate")
While Not t.rs.EOF
    Combo3.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
t.rs.Close
End Sub

Private Sub Combo3_Change()
If Not Combo3.Text = "" Then
    If Combo2.Text = "ID" Or Combo2.Text = "Votes" Then
        t.Database ("SELECT * FROM Candidate WHERE " & Combo2.Text & "=" & Val(Combo3.Text))
    Else
        t.Database ("SELECT * FROM Candidate WHERE " & Combo2.Text & "='" & Combo3.Text & "'")
    End If
    Call datagrid
End If
End Sub

Private Sub Combo3_Click()
If Combo2.Text = "ID" Or Combo2.Text = "Votes" Then
    t.Database ("SELECT * FROM Candidate WHERE " & Combo2.Text & "=" & Val(Combo3.Text))
Else
    t.Database ("SELECT * FROM Candidate WHERE " & Combo2.Text & "='" & Combo3.Text & "'")
End If
Call datagrid
End Sub

Private Sub Command1_Click()
t.Database ("Select * from Post")
t.db.Execute ("Delete from Winners")
t2.Database ("Select Name From Courses")
Set rs = New ADODB.Recordset
Set rs1 = New ADODB.Recordset
While Not t.rs.EOF '1
    If t.rs.Fields(2) <> "Class" Then
    rs.Open "Select * From Candidate where Post='" + t.rs.Fields(1) + "' and Votes IN(Select Max(Votes) from Candidate where Post='" + t.rs.Fields(1) + "')", t.db, adOpenKeyset, adLockOptimistic
    rs1.Open "Select * From Winners", t.db, adOpenKeyset, adLockOptimistic
    While Not rs.EOF '2
        With rs1
            .AddNew
            .Fields(0) = rs.Fields(0)
            .Fields(1) = rs.Fields(2)
            .Fields(2) = rs.Fields(3)
            .Fields(3) = rs.Fields(4)
            .Fields(4) = rs.Fields(6)
            .Fields(5) = rs.Fields(7)
            .Fields(6) = rs.Fields(8)
            .Update
            rs.MoveNext
        End With
    Wend '2/
    rs.Close
    rs1.Close
    Else
    While Not t2.rs.EOF '2
        For i = 0 To 2
            query = "Select * From Candidate where Post='" + t.rs.Fields(1) + "' and Class='" + t2.rs.Fields(0) + "'"
            query = query + " and Year='" + year(i) + "' and Votes IN("
            query = query + "Select Max(Votes) from Candidate where Post='" + t.rs.Fields(1) + "' and Class='" + t2.rs.Fields(0) + "'  and Year='" + year(i) + "')"
            rs.Open query, t.db, adOpenKeyset, adLockOptimistic
            rs1.Open "Select * From Winners", t.db, adOpenKeyset, adLockOptimistic
            While Not rs.EOF '3
                With rs1
                    .AddNew
                    .Fields(0) = rs.Fields(0)
                    .Fields(1) = rs.Fields(2)
                    .Fields(2) = rs.Fields(3)
                    .Fields(3) = rs.Fields(4)
                    .Fields(4) = rs.Fields(6)
                    .Fields(5) = rs.Fields(7)
                    .Fields(6) = rs.Fields(8)
                    .Update
                    rs.MoveNext
                End With
            Wend '3/
            rs.Close
            rs1.Close
        Next
    t2.rs.MoveNext
    Wend '/2
    End If
t.rs.MoveNext
Wend '1
frmwin.Show vbModal, Me
End Sub

Private Sub Command2_Click()
Call Form_Load
End Sub

Private Sub Command3_Click()
Login.lblelection.Caption = "OFF"
t.Database ("Select * from status")
t.rs.MoveFirst
t.rs.Update 0, "OFF"
DataGrid1.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Frame1.Visible = False
Command4.Visible = True
Command5.Visible = True
DataGrid1.Refresh
End Sub

Private Sub Command4_Click()
Login.lblelection.Caption = "ON"
t.Database ("Select * from status")
t.rs.MoveFirst
t.rs.Update 0, "ON"
DataGrid1.Visible = True
Picture1.Visible = True
Picture2.Visible = True
Frame1.Visible = True
Command4.Visible = False
Command5.Visible = False
t.Database ("Select * from Candidate")
t.db.Execute ("delete from Winners")
While Not t.rs.EOF
    t.rs.Fields("Votes") = 0
    t.rs.Update
    t.rs.MoveNext
Wend
End Sub

Private Sub Command5_Click()
Unload Me
frmwin.Show vbModal, MDIForm1
End Sub

Private Sub Form_Load()
t.Database ("Select * from Candidate")
Call datagrid
Combo1.AddItem "ID"
Combo1.AddItem "V_ID"
Combo1.AddItem "Name"
Combo1.AddItem "Class"
Combo1.AddItem "Year"
Combo1.AddItem "Party"
Combo1.AddItem "Post"
Combo1.AddItem "Votes"
Combo2.AddItem "ID"
Combo2.AddItem "V_ID"
Combo2.AddItem "Name"
Combo2.AddItem "Class"
Combo2.AddItem "Year"
Combo2.AddItem "Party"
Combo2.AddItem "Post"
Combo2.AddItem "Votes"
Combo1.ListIndex = 0
Combo2.ListIndex = 0
Option1.Value = False
Option2.Value = False
Combo3.Text = ""
year(0) = "First"
year(1) = "Second"
year(2) = "Third"
If Login.lblelection.Caption = "OFF" Then
    Call Command3_Click
Else
    Call Command4_Click
End If
End Sub

Private Sub Option1_Click()
If Not Combo1.Text = "" Then
t.Database ("SELECT * FROM Candidate ORDER BY " & Combo1.Text & " ASC")
Call datagrid
End If
End Sub

Private Sub Option2_Click()
If Not Combo1.Text = "" Then
t.Database ("SELECT * FROM Candidate ORDER BY " & Combo1.Text & " DESC")
Call datagrid
End If
End Sub

