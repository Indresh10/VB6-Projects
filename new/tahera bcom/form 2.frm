VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form supplier 
   BackColor       =   &H008080FF&
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12840
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   12840
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      Height          =   495
      Left            =   5040
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Exit"
      Height          =   495
      Left            =   7920
      TabIndex        =   10
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Edit"
      Height          =   495
      Left            =   6000
      TabIndex        =   9
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4200
      TabIndex        =   8
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   495
      Left            =   2400
      TabIndex        =   7
      Top             =   4440
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   5880
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   5
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5880
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Width           =   3135
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   2520
      Top             =   3960
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
      RecordSource    =   "supplier"
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   2
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label2 
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
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1680
      TabIndex        =   1
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SUPPLIERS DETAIL"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   360
      Width           =   5895
   End
End
Attribute VB_Name = "supplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim com As Integer

Private Sub Command1_Click()
If com = 1 And Text1.Text <> "" And Text2.Text <> "" And Text3.Text <> "" Then
Adodc1.Recordset.Update
en True
ElseIf com = 2 Then
en True
Else
mesg = MsgBox("Fill All Fields", vbExclamation)
End If
End Sub

Private Sub Command3_Click()
en False
Adodc1.Recordset.AddNew
com = 1
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.Delete
If Adodc1.Recordset.BOF = True Then
 Adodc1.Recordset.MoveLast
Else
 Adodc1.Recordset.MovePrevious
End If
End Sub

Private Sub Command5_Click()
en False
com = 2
End Sub

Private Sub Command6_Click()
Unload Me
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


