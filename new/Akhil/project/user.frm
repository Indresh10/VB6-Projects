VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form student 
   Caption         =   "STUDENT"
   ClientHeight    =   6435
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11385
   BeginProperty Font 
      Name            =   "Viner Hand ITC"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form10"
   Picture         =   "user.frx":0000
   ScaleHeight     =   6435
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   855
      Left            =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1508
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\Akhil\project\res.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\Akhil\project\res.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.PictureBox Picture1 
      BackColor       =   &H8000000B&
      Height          =   615
      Left            =   0
      ScaleHeight     =   555
      ScaleWidth      =   25995
      TabIndex        =   0
      Top             =   0
      Width           =   26055
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   615
         Left            =   17280
         TabIndex        =   6
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   615
         Left            =   5640
         TabIndex        =   5
         Top             =   0
         Width           =   10695
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Label4"
         Height          =   615
         Left            =   1440
         TabIndex        =   4
         Top             =   0
         Width           =   3015
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Roll No."
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   0
         Width           =   1815
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   615
         Left            =   16440
         TabIndex        =   2
         Top             =   0
         Width           =   1335
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   375
         Left            =   4560
         TabIndex        =   1
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Menu detal 
      Caption         =   "DETAILS"
   End
   Begin VB.Menu res 
      Caption         =   "RESULT"
   End
   Begin VB.Menu set 
      Caption         =   "SETTINGS"
      Begin VB.Menu chng 
         Caption         =   "CHANGE PASSWORD"
      End
   End
   Begin VB.Menu ext 
      Caption         =   "EXIT"
   End
End
Attribute VB_Name = "student"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chng_Click()
cngpass.Text4.Text = Label5.Caption
cngpass.Show vbModal, Me
End Sub

Private Sub detal_Click()
Load stddetails
stddetails.Show
Me.Hide
End Sub

Private Sub ext_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label5.Caption = login.usernm.Text
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\res.mdb"
Adodc1.RecordSource = "Select name,rollno,class from student where name='" + Label5.Caption + "'"
Adodc1.Refresh
Label4.Caption = Adodc1.Recordset.Fields(1)
Label5.Caption = Adodc1.Recordset.Fields(0)
Label6.Caption = Adodc1.Recordset.Fields(2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
login.Show
End Sub

Private Sub res_Click()
Unload DataEnvironment1
Load DataEnvironment1
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\res.mdb"
DataEnvironment1.result_Grouping Label6.Caption, CInt(Val(Label4.Caption))
Adodc1.RecordSource = "select name,fathernm,mothernm from student where rollno=" & Val(Label4.Caption) & " and class='" + Label6.Caption + "'"
Adodc1.Refresh
indireport.Sections("Section6").Controls("Label6").Caption = Adodc1.Recordset.Fields(0)
indireport.Sections("Section6").Controls("Label7").Caption = Adodc1.Recordset.Fields(1)
indireport.Sections("Section6").Controls("Label8").Caption = Adodc1.Recordset.Fields(2)
Adodc1.RecordSource = "select stat,per,total from class_res where rollno=" & Val(Label4.Caption) & "and class='" + Label6.Caption + "'"
Adodc1.Refresh
indireport.Sections("Section3").Controls("Label17").Caption = Adodc1.Recordset.Fields(2)
indireport.Sections("Section3").Controls("Label12").Caption = Adodc1.Recordset.Fields(0)
indireport.Sections("Section3").Controls("Label16").Caption = Adodc1.Recordset.Fields(1)
indireport.Show
End Sub
