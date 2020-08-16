VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmvote_party 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Vote To Party"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6885
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmvote_party.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   615
      Left            =   3600
      TabIndex        =   4
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Vote"
      Height          =   615
      Left            =   1200
      TabIndex        =   3
      Top             =   6000
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      BorderStyle     =   0  'None
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6855
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   5055
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8916
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         BorderStyle     =   0
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
      Begin VB.ComboBox Combo1 
         Height          =   660
         Left            =   3360
         TabIndex        =   2
         Top             =   0
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Select Party"
         Height          =   615
         Left            =   1200
         TabIndex        =   1
         Top             =   0
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmvote_party"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1

Private Sub Combo1_Change()
datagrid ("Select * from sub_candidate where Party='" + Combo1.Text + "'")
End Sub

Private Sub Combo1_Click()
datagrid ("Select * from sub_candidate where Party='" + Combo1.Text + "'")
End Sub

Private Sub Command1_Click()
If Trim(Combo1.Text) = "Independent" Then
    MsgBox "Sorry Can't Vote to Independent Candidate", vbCritical
ElseIf Trim(Combo1.Text) = "" Then
    MsgBox "Select party First", vbExclamation
Else
    For i = 0 To 7
        t.Database ("Select Name,Post from sub_candidate where Party='" + Combo1.Text + "' and Post='" + frmvote.Label3(i).Caption + "'")
        If Not t.rs.EOF Then: frmvote.Combo1(i).Text = t.rs.Fields(0)
    Next
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
t.Database ("Select * from Party")
t.rs.MoveFirst
While Not t.rs.EOF
    Combo1.AddItem t.rs.Fields(1)
    t.rs.MoveNext
Wend
t.rs.Close
t.db.Close
datagrid ("Select * from sub_candidate")
End Sub
Private Function datagrid(query As String)
t.Database (query)
Set DataGrid1.DataSource = t.rs
DataGrid1.Columns(0).Width = 500
DataGrid1.Refresh
End Function


