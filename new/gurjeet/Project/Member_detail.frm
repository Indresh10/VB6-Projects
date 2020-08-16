VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMbrDtl 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Member Detail"
   ClientHeight    =   6210
   ClientLeft      =   2460
   ClientTop       =   945
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   1800
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Connect         =   $"Member_detail.frx":0000
      OLEDBString     =   $"Member_detail.frx":008C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Member_Query "
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
   Begin VB.CommandButton Command1 
      Caption         =   "Create Report"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   5280
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "Member_detail.frx":0118
      Height          =   3855
      Left            =   180
      TabIndex        =   8
      Top             =   1200
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   6800
      _Version        =   393216
      DefColWidth     =   80
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
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
            LCID            =   16393
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
            LCID            =   16393
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   660
      Left            =   810
      TabIndex        =   3
      Top             =   420
      Width           =   5415
      Begin VB.ComboBox CmbClassYear 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Member_detail.frx":012D
         Left            =   3840
         List            =   "Member_detail.frx":014C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   1215
      End
      Begin VB.ComboBox CmbClass 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Member_detail.frx":0180
         Left            =   990
         List            =   "Member_detail.frx":019C
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   190
         Width           =   1215
      End
      Begin VB.Label LblYear 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3240
         TabIndex        =   7
         Top             =   250
         Width           =   525
      End
      Begin VB.Label LblClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   270
         TabIndex        =   6
         Top             =   255
         Width           =   600
      End
   End
   Begin VB.CommandButton CmdRef 
      Caption         =   "&Refresh"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   1
      Top             =   5280
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   1800
      Top             =   5280
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
      Connect         =   $"Member_detail.frx":01CF
      OLEDBString     =   $"Member_detail.frx":025B
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Member_Query "
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
   Begin VB.Label LblMbrDtl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Member Detail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   435
      Left            =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2370
   End
End
Attribute VB_Name = "FrmMbrDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmbClass_Click()
Adodc2.RecordSource = "Select Distinct Year From Member_Query where Class='" + CmbClass.Text + "'"
Adodc2.Refresh
CmbClassYear.Clear
While Not Adodc2.Recordset.EOF
CmbClassYear.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Wend
Adodc1.RecordSource = "Select * From Member_Query where class ='" + CmbClass.Text + "'"
Adodc1.Refresh
DataGrid1.Refresh
DataGrid1.Refresh
End Sub

Private Sub CmbClassYear_Click()
Adodc1.RecordSource = "Select * From Member_Query where class ='" + CmbClass.Text + "' and year='" + CmbClassYear.Text + "'"
Adodc1.Refresh
DataGrid1.Refresh
DataGrid1.Refresh
End Sub

Private Sub CmdCancel_Click()
Unload Me
End Sub

Private Sub CmdRef_Click()
CmbClass.Clear
CmbClassYear.Clear
Adodc1.RecordSource = "Select * from Member_Query"
Adodc1.Refresh
DataGrid1.Refresh
DataGrid1.Refresh
Call Form_Load
End Sub

Private Sub Command1_Click()
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
db.Open Adodc1.ConnectionString
rs.Open Adodc1.RecordSource, db, adOpenKeyset, adLockOptimistic
Set MBReport1.DataSource = rs
MBReport1.Show vbModal, Me
End Sub

Private Sub Form_Load()
Adodc2.RecordSource = "Select Distinct Class From Member_Query"
Adodc2.Refresh
CmbClass.Clear
While Not Adodc2.Recordset.EOF
CmbClass.AddItem Adodc2.Recordset.Fields(0)
Adodc2.Recordset.MoveNext
Wend
End Sub
