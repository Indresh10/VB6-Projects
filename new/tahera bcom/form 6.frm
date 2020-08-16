VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form customer 
   BackColor       =   &H00808000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form6"
   ClientHeight    =   8670
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10635
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8670
   ScaleWidth      =   10635
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "save"
      Height          =   495
      Left            =   4320
      TabIndex        =   11
      Top             =   6840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   7320
      TabIndex        =   10
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   5400
      TabIndex        =   9
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "DELETE"
      Height          =   495
      Left            =   3600
      TabIndex        =   8
      Top             =   6840
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD"
      Height          =   495
      Left            =   1800
      TabIndex        =   7
      Top             =   6840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1920
      Top             =   6240
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\user\Desktop\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "customer"
      Caption         =   ""
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
   Begin VB.TextBox Text3 
      DataField       =   "cmob"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   4560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      DataField       =   "caddress"
      DataSource      =   "Adodc1"
      Height          =   855
      Left            =   5280
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   3240
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      DataField       =   "cname"
      DataSource      =   "Adodc1"
      Height          =   495
      Left            =   5280
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2160
      Width           =   4095
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "MOB.NO."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   3
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   3480
      Width           =   1815
   End
   Begin VB.Label TGTETGERTGR 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "CUSTOMERS DETAIL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   6135
   End
End
Attribute VB_Name = "customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim com As Integer
Private Sub Command1_Click()
en False
Adodc1.Recordset.AddNew
com = 1
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.Delete
If Adodc1.Recordset.BOF = True Then
 Adodc1.Recordset.MoveLast
Else
 Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command3_Click()
en False
com = 2
End Sub

Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
If com = 1 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
Adodc1.Recordset.Update
en True
ElseIf com = 2 Then
en True
Else
mesg = MsgBox("Fill All Fields", vbExclamation)
End If
End Sub

Private Sub Form_Load()
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Fruit.mdb;Persist Security Info=False"
End Sub

Private Function en(a As Boolean)
Text1.Locked = a
Text2.Locked = a
Text3.Locked = a
Command1.Visible = a
Command2.Visible = a
Command3.Visible = a
Command4.Visible = a
If a = True Then
 Command5.Visible = False
Else
 Command5.Visible = True
End If
End Function

