VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmwin 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Winners"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   11295
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmwin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   11295
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFA0&
      Caption         =   "Print Winners"
      Height          =   615
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5760
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00FFFFA0&
      Caption         =   "Print Preview"
      Height          =   615
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2535
   End
   Begin MSFlexGridLib.MSFlexGrid fg1 
      Height          =   4215
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   7435
      _Version        =   393216
      Cols            =   7
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16777152
      BackColorSel    =   8421631
      BackColorBkg    =   12648447
      Enabled         =   -1  'True
      TextStyle       =   4
      TextStyleFixed  =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe Print"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Winners As Of"
      Height          =   540
      Left            =   4500
      TabIndex        =   4
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Winners As Of"
      Height          =   615
      Left            =   4440
      TabIndex        =   1
      Top             =   120
      Width           =   2415
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmwin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1

Private Sub Command1_Click()
t.Database ("Select * from Winners")
Set winreport.DataSource = t.rs
winreport.Show vbModal, Me
End Sub

Private Sub Command2_Click()
t.Database ("Select * from Winners")
Set winreport.DataSource = t.rs
winreport.PrintReport True
End Sub

Private Sub Form_Load()
Label2.Caption = Format(Now, "dd-mmm-yyyy hh:mm:ss")
fg1.TextMatrix(0, 0) = "ID"
fg1.TextMatrix(0, 1) = "Name"
fg1.TextMatrix(0, 2) = "Class"
fg1.TextMatrix(0, 3) = "Year"
fg1.TextMatrix(0, 4) = "Party"
fg1.TextMatrix(0, 5) = "Post"
fg1.TextMatrix(0, 6) = "Votes"
t.Database ("Select * From Winners order by C_ID")
t.rs.MoveFirst
While Not t.rs.EOF
    For i = 0 To 6
        fg1.TextMatrix(fg1.Rows - 1, i) = t.rs.Fields(i)
    Next
    t.rs.MoveNext
    If Not t.rs.EOF Then: fg1.Rows = fg1.Rows + 1
Wend
End Sub
