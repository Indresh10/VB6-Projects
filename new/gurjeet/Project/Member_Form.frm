VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form FrmMember 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Member Operations"
   ClientHeight    =   7695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   17175
   Icon            =   "Member_Form.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7695
   ScaleWidth      =   17175
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text3 
      DataField       =   "DOJ"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   3600
      MaxLength       =   20
      TabIndex        =   40
      Top             =   1800
      Visible         =   0   'False
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1320
      Top             =   6120
      Width           =   3855
      _ExtentX        =   6800
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
      Appearance      =   0
      BackColor       =   12640511
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Member_Form.frx":08CA
      OLEDBString     =   $"Member_Form.frx":0956
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from member"
      Caption         =   "                    MEMBER"
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
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1200
      TabIndex        =   35
      Top             =   1800
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   114884611
      CurrentDate     =   43837
   End
   Begin VB.Frame FremCategory 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Search Category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6660
      Left            =   6345
      TabIndex        =   31
      Top             =   720
      Width           =   10830
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Code"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3960
         TabIndex        =   44
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   6240
         TabIndex        =   43
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   5040
         TabIndex        =   42
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Year"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   7200
         TabIndex        =   41
         Top             =   360
         Width           =   1215
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Member_Form.frx":09E2
         Height          =   5295
         Left            =   120
         TabIndex        =   38
         Top             =   1200
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   9340
         _Version        =   393216
         AllowUpdate     =   0   'False
         DefColWidth     =   67
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
      Begin VB.TextBox TxtSearch 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   5550
         MaxLength       =   15
         TabIndex        =   33
         Top             =   720
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Searching word :"
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
         Left            =   3960
         TabIndex        =   32
         Top             =   780
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtLast 
      DataField       =   "father"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   2970
      MaxLength       =   15
      TabIndex        =   12
      Top             =   1320
      Width           =   1845
   End
   Begin VB.TextBox TxtFirst 
      DataField       =   "Name"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   11
      Top             =   1320
      Width           =   1770
   End
   Begin VB.TextBox TxtCity 
      DataField       =   "City"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   16
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox TxtAddress 
      DataField       =   "Address"
      DataSource      =   "Adodc1"
      Height          =   1215
      Left            =   1200
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   2280
      Width           =   4935
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   630
      Left            =   735
      Picture         =   "Member_Form.frx":09F7
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Add Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   630
      Left            =   1575
      Picture         =   "Member_Form.frx":0D39
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Edit Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   630
      Left            =   4095
      Picture         =   "Member_Form.frx":107B
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Search Record"
      Top             =   6735
      Width           =   735
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "&Delete"
      Height          =   630
      Left            =   2415
      Picture         =   "Member_Form.frx":13BD
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Delete Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   630
      Left            =   3255
      Picture         =   "Member_Form.frx":16FF
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Save Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   630
      Left            =   4935
      Picture         =   "Member_Form.frx":1D69
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Exit"
      Top             =   6750
      Width           =   735
   End
   Begin VB.Frame FremPerInfo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Personal Information"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   240
      TabIndex        =   19
      Top             =   4680
      Width           =   5895
      Begin VB.TextBox Txtgender 
         DataField       =   "Gender"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3720
         TabIndex        =   39
         Top             =   720
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.OptionButton OptFemale 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Female"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         MaskColor       =   &H00C0FFC0&
         TabIndex        =   24
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton OptMale 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Male"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   23
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox TxtContact 
         Alignment       =   1  'Right Justify
         DataField       =   "Cnt_No"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   21
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label LblGender 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Gender :"
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
         Left            =   3720
         TabIndex        =   22
         Top             =   420
         Width           =   765
      End
      Begin VB.Label LblContact 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
         Caption         =   "Contact No :"
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
         Left            =   240
         TabIndex        =   20
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.TextBox TxtFee 
      Alignment       =   1  'Right Justify
      DataField       =   "Fee"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   18
      Top             =   3600
      Width           =   1215
   End
   Begin VB.TextBox TxtSurname 
      DataField       =   "surname"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4800
      MaxLength       =   15
      TabIndex        =   10
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox TxtCode 
      Alignment       =   1  'Right Justify
      DataField       =   "Code"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   675
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   660
      Left            =   600
      TabIndex        =   1
      Top             =   3960
      Width           =   5415
      Begin VB.TextBox Text2 
         DataField       =   "Year"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   3840
         TabIndex        =   37
         Top             =   200
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         DataField       =   "Class"
         DataSource      =   "Adodc1"
         Height          =   375
         Left            =   960
         TabIndex        =   36
         Top             =   200
         Width           =   1935
      End
      Begin VB.Label LblClass 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   2
         Top             =   255
         Width           =   600
      End
      Begin VB.Label LblYear 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0E0FF&
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
         TabIndex        =   3
         Top             =   250
         Width           =   525
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   255
      Left            =   6720
      TabIndex        =   45
      Top             =   960
      Width           =   615
   End
   Begin VB.Label LblJoinDate 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Join Date :"
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
      Left            =   240
      TabIndex        =   34
      Top             =   1860
      Width           =   945
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBER OPERATIONS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   6083
      TabIndex        =   0
      Top             =   45
      Width           =   4125
   End
   Begin VB.Shape ShapLabel 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   17160
   End
   Begin VB.Label LblLast 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Father Name"
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
      Left            =   3360
      TabIndex        =   8
      Top             =   1080
      Width           =   1170
   End
   Begin VB.Label LblFirst 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Member Name"
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
      Left            =   1440
      TabIndex        =   7
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label LblSurname 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Surname"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   1080
      Width           =   810
   End
   Begin VB.Label LblFee 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Member Fee :"
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
      Left            =   3600
      TabIndex        =   17
      Top             =   3660
      Width           =   1245
   End
   Begin VB.Label LblCity 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "City :"
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
      Left            =   240
      TabIndex        =   15
      Top             =   3660
      Width           =   420
   End
   Begin VB.Label LblAddress 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Address :"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2310
      Width           =   855
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Name :"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1380
      Width           =   645
   End
   Begin VB.Label LblCode 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Code :"
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
      Left            =   240
      TabIndex        =   4
      Top             =   735
      Width           =   585
   End
End
Attribute VB_Name = "FrmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub CmdAdd_Click()
Adodc1.Recordset.AddNew
TxtFee.Text = Label2.Caption
tenable (True)
cenable (False)
End Sub

Private Sub CmdCancel_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
tenable (False)
cenable (True)
Call refr
End Sub

Private Sub CmdDel_Click()
If Adodc1.Recordset.BOF = False And Adodc1.Recordset.EOF = False Then
    Adodc1.Recordset.Delete
    Adodc1.Recordset.MovePrevious
    If Adodc1.Recordset.BOF Then Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub CmdEdit_Click()
tenable (True)
cenable (False)
CmdCancel.Enabled = False
End Sub

Private Sub CmdExit_Click()
Unload Me
End Sub

Private Sub CmdSave_Click()
If TxtCode.Text = "" Or TxtFirst.Text = "" Or TxtLast.Text = "" Or TxtSurname.Text = "" Or TxtAddress.Text = "" Or TxtCity.Text = "" Or Text1.Text = "" Or TxtContact.Text = "" Or Txtgender.Text = "" Then
MsgBox "please fill all the details", vbExclamation
Exit Sub
Else
tenable (False)
cenable (True)
Adodc1.Recordset.Update
Call refr
Call refr
End If
End Sub

Private Sub DTPicker1_Change()
Text3.Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
Label2.Caption = MDIFrm.Label15.Caption
tenable (False)
cenable (True)
DataGrid1.Columns(7).Width = 500
DataGrid1.Columns(8).Width = 500
DataGrid1.Columns(10).Width = 500
DataGrid1.Columns(11).Width = 500
DataGrid1.Columns(12).Visible = False
End Sub

Private Sub Opt1_Click(Index As Integer)
TxtSearch.Enabled = True
TxtSearch.SetFocus
i = Index
End Sub

Private Sub OptFemale_Click()
Txtgender.Text = "F"
End Sub

Private Sub OptMale_Click()
Txtgender.Text = "M"
End Sub

Private Sub Text3_Change()
If Not Text3.Text = "" Then DTPicker1.Value = Text3.Text
End Sub

Private Sub Txtgender_Change()
If Txtgender.Text = "M" Then
    OptMale.Value = True
ElseIf Txtgender.Text = "F" Then
    OptFemale.Value = True
Else
    OptMale.Value = False
    OptFemale.Value = False
End If
End Sub
Private Function tenable(a As Boolean)
TxtCode.Enabled = a
TxtFirst.Enabled = a
TxtLast.Enabled = a
TxtSurname.Enabled = a
TxtAddress.Enabled = a
TxtCity.Enabled = a
TxtFee.Enabled = False
Text1.Enabled = a
Text2.Enabled = a
TxtContact.Enabled = a
OptMale.Enabled = a
OptFemale.Enabled = a
DTPicker1.Enabled = a
End Function

Private Function cenable(a As Boolean)
CmdAdd.Enabled = a
CmdEdit.Enabled = a
CmdDel.Enabled = a
CmdSave.Enabled = Not a
CmdCancel.Enabled = Not a
End Function
Private Function refr()
Adodc1.Refresh
DataGrid1.Refresh
DataGrid1.Columns(7).Width = 500
DataGrid1.Columns(8).Width = 500
DataGrid1.Columns(10).Width = 500
DataGrid1.Columns(11).Width = 500
DataGrid1.Columns(12).Visible = False
End Function

Private Sub TxtSearch_Change()
Adodc1.RecordSource = "select * from Member where " + Opt1(i).Caption + " like '" + TxtSearch.Text + "%'"
Call refr
Call refr
End Sub
