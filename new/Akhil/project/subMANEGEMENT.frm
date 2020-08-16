VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form SUBJECT 
   Caption         =   "SUBJECTS"
   ClientHeight    =   7905
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14865
   BeginProperty Font 
      Name            =   "Viner Hand ITC"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "subMANEGEMENT.frx":0000
   ScaleHeight     =   7905
   ScaleWidth      =   14865
   Begin VB.Frame Frame1 
      Caption         =   "Search"
      Height          =   2775
      Left            =   3360
      TabIndex        =   15
      Top             =   2040
      Visible         =   0   'False
      Width           =   8295
      Begin VB.CommandButton Command7 
         Caption         =   "Ok"
         Height          =   870
         Left            =   0
         TabIndex        =   18
         Top             =   1920
         Width           =   8295
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "subjectname"
         DataSource      =   "Adodc1"
         Height          =   630
         Left            =   840
         TabIndex        =   17
         Text            =   "Combo1"
         Top             =   1200
         Width           =   6615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   " Subject name"
         Height          =   510
         Left            =   2880
         TabIndex        =   16
         Top             =   480
         Width           =   2745
      End
   End
   Begin VB.CommandButton Command6 
      Caption         =   "SEARCH"
      Height          =   750
      Left            =   9360
      TabIndex        =   14
      Top             =   5280
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "CANCEL"
      Height          =   750
      Left            =   7680
      TabIndex        =   13
      Top             =   6500
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "DELETE"
      Height          =   750
      Left            =   7320
      TabIndex        =   12
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      DataField       =   "max"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   630
      Left            =   10200
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "min"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   630
      Left            =   6720
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "subjectname"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   630
      Left            =   6720
      TabIndex        =   2
      Top             =   3120
      Width           =   4695
   End
   Begin VB.TextBox Text1 
      DataField       =   "subjectcode"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   630
      Left            =   6720
      TabIndex        =   1
      Top             =   2040
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   570
      Left            =   3240
      Top             =   6000
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   1005
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"subMANEGEMENT.frx":67185
      OLEDBString     =   $"subMANEGEMENT.frx":6720C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from subject"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "SAVE"
      Height          =   750
      Left            =   3240
      TabIndex        =   6
      Top             =   6500
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      Height          =   750
      Left            =   5280
      TabIndex        =   5
      Top             =   5280
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   750
      Left            =   3240
      TabIndex        =   0
      Top             =   5280
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      Height          =   510
      Left            =   6720
      TabIndex        =   19
      Top             =   1200
      Width           =   1065
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Min"
      Height          =   510
      Left            =   4920
      TabIndex        =   11
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Max"
      Height          =   510
      Left            =   8760
      TabIndex        =   10
      Top             =   4200
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subject name"
      Height          =   510
      Left            =   3480
      TabIndex        =   9
      Top             =   3240
      Width           =   2175
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Subject code"
      ForeColor       =   &H00000000&
      Height          =   510
      Left            =   3480
      TabIndex        =   8
      Top             =   2160
      Width           =   1920
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Management Of Subjects"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4920
      TabIndex        =   7
      Top             =   480
      Width           =   5175
   End
End
Attribute VB_Name = "SUBJECT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.AddNew
tenable (True)
chide (False)
End Sub

Private Sub Command2_Click()
tenable (True)
chide (False)
Command5.Visible = False
End Sub

Private Sub Command3_Click()
tenable (False)
chide (True)
Adodc1.Recordset.Fields("class") = Label3.Caption
Adodc1.Recordset.Update
End Sub

Private Sub Command4_Click()
If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command5_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
tenable False
chide True
End Sub

Private Sub Command6_Click()
Combo1.Clear
Combo1.AddItem ""
Adodc1.Recordset.MoveFirst
While Not Adodc1.Recordset.EOF
    Combo1.AddItem Adodc1.Recordset.Fields("subjectname")
    Adodc1.Recordset.MoveNext
Wend
Adodc1.Recordset.MoveFirst
Adodc1.Refresh
Frame1.Visible = True
End Sub

Private Sub Command7_Click()
If Combo1.Text <> "" Then
Adodc1.Recordset.Find ("subjectname='" + Combo1.Text + "'")
Else
MsgBox "Please Select Name of subject", vbExclamation
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\res.mdb"
Adodc1.Refresh
Label3.Caption = teach.Label6.Caption
Adodc1.RecordSource = "select * from subject where class='" + Label3.Caption + "'"
Adodc1.Refresh
End Sub
Private Function tenable(a As Boolean)
 Text1.Enabled = a
 Text2.Enabled = a
 Text3.Enabled = a
 Text4.Enabled = a
End Function
Private Function chide(b As Boolean)
Command1.Visible = b
Command2.Visible = b
Command4.Visible = b
Command6.Visible = b
Command5.Visible = Not b
Command3.Visible = Not b
End Function

