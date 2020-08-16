VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmclg_data 
   BackColor       =   &H00AAFFAA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Data From College Database "
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   16020
   BeginProperty Font 
      Name            =   "Lucida Calligraphy"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmclg_data.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   16020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Refresh"
      Height          =   855
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8040
      Width           =   2175
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3735
      Left            =   13080
      ScaleHeight     =   3705
      ScaleWidth      =   2745
      TabIndex        =   10
      Top             =   5160
      Width           =   2775
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFF80&
         Caption         =   "Add to Database"
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00FFFF80&
         Caption         =   "Generate Voter ID"
         Height          =   855
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1680
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFF80&
         Caption         =   "Import Data from file"
         Height          =   855
         Left            =   120
         MaskColor       =   &H00000000&
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   735
         Left            =   120
         TabIndex        =   16
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00AFFFF0&
         Caption         =   "Options"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   13080
      ScaleHeight     =   2505
      ScaleWidth      =   2745
      TabIndex        =   4
      Top             =   120
      Width           =   2775
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
         TabIndex        =   7
         Top             =   720
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
         TabIndex        =   6
         Top             =   1320
         UseMaskColor    =   -1  'True
         Width           =   1815
      End
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
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00AFFFF0&
         Caption         =   "Sort By"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   2535
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   13080
      ScaleHeight     =   2265
      ScaleWidth      =   2745
      TabIndex        =   0
      Top             =   2760
      Width           =   2775
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
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
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
         TabIndex        =   1
         Top             =   1440
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
         TabIndex        =   3
         Top             =   120
         Width           =   2535
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   8775
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   15478
      _Version        =   393216
      AllowUpdate     =   0   'False
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   7800
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmclg_data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New DataSource
Dim t1 As New Database1
Dim db As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As New ADODB.Recordset
Private Sub Combo1_Click()
Option1.Value = False
Option2.Value = False
End Sub
Private Sub Combo2_Click()
Combo3.Text = ""
End Sub
Private Sub Combo3_Change()
If Not Combo3.Text = "" Then
    If Not Combo2.Text = "DOB" Then
        Call datagrid("Select * from Stud_Dtls where " + Combo2.Text + "='" + Combo3.Text + "'")
    Else
        Call datagrid("Select * from Stud_Dtls where " + Combo2.Text + "=#" + Combo3.Text + "#")
    End If
End If
End Sub

Private Sub Combo3_Click()
If Not Combo2.Text = "DOB" Then
    Call datagrid("Select * from Stud_Dtls where " + Combo2.Text + "='" + Combo3.Text + "'")
Else
    Call datagrid("Select * from Stud_Dtls where " + Combo2.Text + "=#" + Combo3.Text + "#")
End If
End Sub

Private Sub Combo3_GotFocus()
Call t.Database(App.Path & "\college.mdb", "Select Distinct " + Combo2.Text + " From Stud_Dtls")
Combo3.clear
While Not t.rs.EOF
    Combo3.AddItem t.rs.Fields(0)
    t.rs.MoveNext
Wend
End Sub

Private Sub Command1_Click()
On Error GoTo cdhandler
start:
With CD
    .Filter = "Microsoft Accesss Document(*.mdb)|*.mdb|Microsoft Excel Document(*.xls)|*.xls"
    .ShowOpen
End With
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
db.Provider = "Microsoft.Jet.OLEDB.4.0"
If Right(CD.FileName, 4) = ".xls" Then
db.ConnectionString = "Data Source = " & CD.FileName & ";" & "Extended Properties=Excel 8.0;"
db.Open
ElseIf Right(CD.FileName, 4) = ".mdb" Then
db.ConnectionString = "Data Source = " & CD.FileName
db.Open
Else
Exit Sub
End If
Set rs1 = db.OpenSchema(adSchemaTables)
While Not rs1.EOF
    Str1 = UCase(Left(rs1.Fields("TABLE_NAME"), 4))
    If Not Str1 = "MSYS" Then frmtable.Combo1.AddItem rs1.Fields("TABLE_NAME")
    rs1.MoveNext
Wend
frmtable.Show vbModal, Me
If Right(CD.FileName, 4) = ".mdb" Then
rs.Open "SELECT * FROM " & Label4.Caption, db, adOpenKeyset, adLockOptimistic
ElseIf Right(CD.FileName, 4) = ".xls" Then
rs.Open "SELECT * FROM [" & Label4.Caption & "]", db, adOpenKeyset, adLockOptimistic
End If
f = 0
For i = 0 To 6
If rs.Fields(i).name = t.rs.Fields(i + 1).name Then f = f + 1
Next
If f = 7 Then
t.db.Execute ("delete from Stud_Dtls")
t.rs.Resync
Set DataGrid1.DataSource = Nothing
DataGrid1.Refresh
id = 1
While Not rs.EOF
t.rs.AddNew
t.rs.Fields(0) = id
For i = 0 To 6
    t.rs.Fields(i + 1) = rs.Fields(i)
Next
t.rs.Update
rs.MoveNext
id = id + 1
Wend
MsgBox "succesfully imported data", vbInformation
Else
MsgBox "please check your table fields in your excel workbook/access document"
End If
Call Form_Load
db.Close
rs.Close
rs1.Close
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
On Error GoTo handle
start:
Call t.Database(App.Path & "\college.mdb", "Select * From Stud_dtls")
t.db.Execute ("Alter Table Stud_Dtls ADD V_ID VARCHAR(255)")
t.rs.MoveFirst
While Not t.rs.EOF
t.rs.Fields("V_ID") = v_id(t.rs.Fields("Name"), t.rs.Fields("Class"), t.rs.Fields("Year"), t.rs.Fields("DOB"), t.rs.Fields("ID"))
t.rs.MoveNext
Wend
If t.rs.EOF Then t.rs.MoveLast: t.rs.Update
Call datagrid("Select * from Stud_Dtls")
MsgBox "Every Student Voter ID Is Generated", vbInformation
Exit Sub
handle:
If Str(Err.Number) = -2147217887 Then Resume Next
Select Case MsgBox(Error(Err.Number) + vbCr + "Error number->" & Str(Err.Number), vbCritical + vbAbortRetryIgnore, "Error")
Case vbAbort
   Resume exitline
Case vbRetry
   Resume start
Case vbIgnore
   Resume Next
End Select
exitline:
End Sub
Private Function v_id(ByVal name, ByVal Class, ByVal year, ByVal dob, ByVal id) As String
Dim vid As String
vid = "CC10"
vid = vid + Left$(name, 2)
t1.Database ("select ID from Courses where Name='" + Class + "'")
vid = vid & t1.rs.Fields(0)
t1.rs.Close
t1.db.Close
Select Case year
    Case "First"
        vid = vid + "01"
    Case "Second"
        vid = vid + "02"
    Case "Third"
        vid = vid + "03"
End Select
vid = vid & id
v_id = UCase(vid)
End Function

Public Function datagrid(query)
Call t.Database(App.Path & "\college.mdb", query)
Set DataGrid1.DataSource = t.rs
DataGrid1.Columns(0).Visible = False
End Function

Private Sub Command3_Click()
MsgBox "Note:The Voter List Will Be Empty" + vbCr + "For Inserting Data", vbInformation
Call t.Database(App.Path & "\college.mdb", "Select * from Stud_Dtls")
t1.Database ("Select * from Voter")
t1.db.Execute ("Delete from Voter")
MsgBox "Only The student with status as 'pass' is added to voter list"
While Not t.rs.EOF
    If UCase(t.rs.Fields("status")) = "PASS" Then
    With t1.rs
        pass = Left$(t.rs.Fields("Name"), 2) + Right$(t.rs.Fields("Adm_no"), 4)
        .AddNew
        .Fields("ID") = t.rs.Fields("ID")
        .Fields("Name") = t.rs.Fields("Name")
        .Fields("V_ID") = t.rs.Fields("V_ID")
        .Fields("Class") = t.rs.Fields("Class")
        .Fields("Year") = t.rs.Fields("Year")
        .Fields("DOB") = t.rs.Fields("DOB")
        .Fields("Password") = pass
        .Fields("Voted") = "No"
        .Fields("Default") = pass
        .Update
    End With
    End If
    t.rs.MoveNext
Wend
t1.rs.Close
t.rs.Close
t.db.Close
t1.db.Close
MsgBox "Data Successfully Added To Voter List", vbInformation
End Sub

Private Sub Command4_Click()
Call Form_Load
End Sub

Private Sub Form_Load()
Call datagrid("Select * from Stud_Dtls")
Combo1.clear
Combo2.clear
Combo1.AddItem "Adm_no"
Combo1.AddItem "Name"
Combo1.AddItem "Class"
Combo1.AddItem "Year"
Combo1.AddItem "DOB"
Combo1.AddItem "V_ID"
Combo1.AddItem "status"
Combo2.AddItem "Adm_no"
Combo2.AddItem "Name"
Combo2.AddItem "Class"
Combo2.AddItem "Year"
Combo2.AddItem "DOB"
Combo2.AddItem "V_ID"
Combo2.AddItem "status"
Combo3.Text = ""
Combo3.clear
Option1.Value = False
Option2.Value = False
Combo1.ListIndex = 0
Combo2.ListIndex = 0
End Sub

Private Sub Option1_Click()
Call datagrid("Select * From Stud_Dtls Order by " + Combo1.Text + " ASC")
End Sub

Private Sub Option2_Click()
Call datagrid("Select * From Stud_Dtls Order by " + Combo1.Text + " DESC")
End Sub
