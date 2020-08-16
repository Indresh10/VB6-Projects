VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00808000&
   Caption         =   "Form4"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10065
   LinkTopic       =   "Form4"
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   10065
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   2040
      Top             =   6240
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
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
      Connect         =   $"Frmlogin5.frx":0000
      OLEDBString     =   $"Frmlogin5.frx":0099
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "BRF"
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   2040
      Top             =   1560
      Visible         =   0   'False
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   873
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
      Connect         =   $"Frmlogin5.frx":0132
      OLEDBString     =   $"Frmlogin5.frx":01CB
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from bif where ST='Not Returned'"
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
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4920
      TabIndex        =   10
      Text            =   "Combo1"
      Top             =   2160
      Width           =   4095
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   4920
      TabIndex        =   9
      Top             =   4480
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   873
      _Version        =   393216
      Format          =   119275521
      CurrentDate     =   43851
   End
   Begin VB.CommandButton Command3 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7440
      TabIndex        =   8
      Top             =   6720
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "RETURN BOOK"
      BeginProperty Font 
         Name            =   "Elephant"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      TabIndex        =   7
      Top             =   6720
      Width           =   2535
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   4920
      TabIndex        =   6
      Top             =   5640
      Width           =   4095
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   4920
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   3320
      Width           =   4095
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FINES COLLECTED"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   5640
      Width           =   2655
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RETURN DATE"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1320
      TabIndex        =   3
      Top             =   4480
      Width           =   2655
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STUDENT CODE"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOOK ID"
      BeginProperty Font 
         Name            =   "Gloucester MT Extra Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   495
      Left            =   1320
      TabIndex        =   1
      Top             =   3320
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "BOOK RETURN FORM"
      BeginProperty Font 
         Name            =   "Britannic Bold"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2400
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("SC='" + Combo1.Text + "'")
Text2.Text = Adodc1.Recordset.Fields(2)
End If
End Sub

Private Sub Combo1_Click()
If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("SC='" + Combo1.Text + "'")
Text2.Text = Adodc1.Recordset.Fields(2)
End If
End Sub

Private Sub Command1_Click()
Adodc1.Recordset.Update 5, "Returned"
Adodc1.Refresh
Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Trim(Combo1.Text)
Adodc2.Recordset.Fields(1) = Trim(Text2.Text)
Adodc2.Recordset.Fields(2) = Trim(DTPicker1.Value)
Adodc2.Recordset.Fields(3) = Trim(Text3.Text)
Adodc2.Recordset.Update
Adodc2.Refresh
MsgBox "book returned succesfully"
Combo1.Text = ""
Text2.Text = ""
Text3.Text = ""
Call Form_Load
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Form_Load()
DTPicker1.Value = Format(Now, "dd-MM-yyyy")
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Combo1.Clear
While Not Adodc1.Recordset.EOF
    Combo1.AddItem Adodc1.Recordset.Fields(0)
    Adodc1.Recordset.MoveNext
Wend
End Sub
