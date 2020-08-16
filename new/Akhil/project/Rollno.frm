VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Rollno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SELECT ROLL NO."
   ClientHeight    =   3015
   ClientLeft      =   6945
   ClientTop       =   4545
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2400
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   600
      Top             =   1920
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   582
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
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   893
      TabIndex        =   2
      Top             =   1200
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   1133
      TabIndex        =   1
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Select The Roll no.-"
      BeginProperty Font 
         Name            =   "Viner Hand ITC"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   3495
   End
End
Attribute VB_Name = "Rollno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
If Combo1.Text <> "" Then
If Text1.Text = "Mark" Then
markent.Label12.Caption = Combo1.Text
Unload Me
markent.Show
Else
Unload DataEnvironment1
Load DataEnvironment1
DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\res.mdb"
DataEnvironment1.result_Grouping teach.Label6.Caption, CInt(Val(Combo1.Text))
Adodc1.RecordSource = "select name,fathernm,mothernm from student where rollno=" & Val(Combo1.Text) & " and class='" + teach.Label6.Caption + "'"
Adodc1.Refresh
indireport.Sections("Section6").Controls("Label6").Caption = Adodc1.Recordset.Fields(0)
indireport.Sections("Section6").Controls("Label7").Caption = Adodc1.Recordset.Fields(1)
indireport.Sections("Section6").Controls("Label8").Caption = Adodc1.Recordset.Fields(2)
Adodc1.RecordSource = "select stat,per,total from class_res where rollno=" & Val(Combo1.Text) & "and class='" + teach.Label6.Caption + "'"
Adodc1.Refresh
indireport.Sections("Section3").Controls("Label17").Caption = Adodc1.Recordset.Fields(2)
indireport.Sections("Section3").Controls("Label12").Caption = Adodc1.Recordset.Fields(0)
indireport.Sections("Section3").Controls("Label16").Caption = Adodc1.Recordset.Fields(1)
Unload Me
indireport.Show
End If
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\res.mdb"
Adodc1.RecordSource = "select rollno from student where class='" + teach.Label6.Caption + "'"
Adodc1.Refresh
While Not Adodc1.Recordset.EOF
Combo1.AddItem Adodc1.Recordset.Fields(0)
Adodc1.Recordset.MoveNext
Wend
End Sub
