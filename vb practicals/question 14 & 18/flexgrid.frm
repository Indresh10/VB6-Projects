VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form flexgrid 
   Caption         =   "Form1"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6630
   ScaleWidth      =   11670
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid fg1 
      Height          =   2775
      Left            =   3600
      TabIndex        =   2
      Top             =   1800
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   4895
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
   End
   Begin VB.ComboBox Combo1 
      Height          =   495
      ItemData        =   "flexgrid.frx":0000
      Left            =   6120
      List            =   "flexgrid.frx":0002
      TabIndex        =   0
      Text            =   " "
      Top             =   840
      Width           =   2415
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select class"
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   840
      Width           =   1770
   End
End
Attribute VB_Name = "flexgrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As Database
Dim rs As Recordset
Private Sub Combo1_Change()
Set rs = db.OpenRecordset("Select * from new where class='" + Combo1.Text + "'")
r = 1
fg1.TextMatrix(0, 0) = "Name"
fg1.TextMatrix(0, 1) = "Class"
fg1.TextMatrix(0, 2) = "Percentage"
Do Until rs.EOF
    fg1.Rows = r + 1
    For i = 0 To 2
        fg1.TextMatrix(r, i) = rs.Fields(i)
    Next
    r = r + 1
    rs.MoveNext
Loop
End Sub
Private Sub Combo1_Click()
Set rs = db.OpenRecordset("Select * from new where class='" + Combo1.Text + "'")
r = 1
fg1.TextMatrix(0, 0) = "Name"
fg1.TextMatrix(0, 1) = "Class"
fg1.TextMatrix(0, 2) = "Percentage"
Do Until rs.EOF
    fg1.Rows = r + 1
    For i = 0 To 2
        fg1.TextMatrix(r, i) = rs.Fields(i)
    Next
    r = r + 1
    rs.MoveNext
Loop
End Sub

Private Sub Form_Load()
Set db = OpenDatabase("mydb.mdb")
Set rs = db.OpenRecordset("Select distinct class from new")
Do Until rs.EOF
    Combo1.AddItem rs.Fields(0)
    rs.MoveNext
Loop
End Sub
