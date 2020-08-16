VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmdate 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4500
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Show Report"
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   1680
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Select Time period"
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4335
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   375
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   7208961
         CurrentDate     =   43838
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   600
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   7208961
         CurrentDate     =   43838
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   195
         Left            =   2775
         TabIndex        =   4
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   195
         Left            =   840
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2250
      TabIndex        =   5
      Top             =   120
      Width           =   75
   End
End
Attribute VB_Name = "frmdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
Private Sub Command1_Click()
sdate = Format(DTPicker1.Value - 1, "mm-dd-yyyy")
edate = Format(DTPicker2.Value + 1, "mm-dd-yyyy")
Set db = New ADODB.Connection
Set rs = New ADODB.Recordset
If Label3.Caption = "Issue" Then
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\gurjeet\Project\Library.mdb;Persist Security Info=False"
rs.Open "Select * from Issue where Isu_Dt Between #" + sdate + "# and #" + edate + "#", db, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
Set ISReport1.DataSource = rs
Unload Me
ISReport1.Show
Unload Me
Else
MsgBox "No Data Found", vbExclamation
End If
Else
db.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\gurjeet\Project\Library.mdb;Persist Security Info=False"
rs.Open "Select * from Fine where Fin_Dt Between #" + sdate + "# and #" + edate + "#", db, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
Set FNReport1.DataSource = rs
Unload Me
FNReport1.Show
Unload Me
Else
MsgBox "No Data Found", vbExclamation
End If
End If
Set db = Nothing
Set rs = Nothing
End Sub


Private Sub DTPicker1_Change()
DTPicker2.MinDate = DTPicker1.Value
End Sub

Private Sub Form_Load()
DTPicker2.MinDate = DTPicker1.Value
End Sub
