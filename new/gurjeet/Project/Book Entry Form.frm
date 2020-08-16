VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmBkEntry 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book Entry"
   ClientHeight    =   6570
   ClientLeft      =   4260
   ClientTop       =   3015
   ClientWidth     =   13935
   Icon            =   "Book Entry Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   13935
   Begin VB.CommandButton CmdCancel 
      Caption         =   "&Cancel"
      Height          =   630
      Left            =   4260
      Picture         =   "Book Entry Form.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Search Record"
      Top             =   5280
      Width           =   735
   End
   Begin VB.TextBox Text1 
      DataField       =   "DOP"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   4200
      TabIndex        =   28
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   960
      Top             =   4560
      Width           =   4815
      _ExtentX        =   8493
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
      ForeColor       =   0
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Book Entry Form.frx":064C
      OLEDBString     =   $"Book Entry Form.frx":06D8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from book"
      Caption         =   "                                 BOOKS"
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
      Left            =   1680
      TabIndex        =   27
      Top             =   2880
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "dd-MMM-yyyy"
      Format          =   114884611
      CurrentDate     =   43836
   End
   Begin VB.Frame FremCategory 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search &category"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   6720
      TabIndex        =   18
      Top             =   600
      Width           =   7095
      Begin VB.TextBox TxtSearch 
         BackColor       =   &H00C0E0FF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   3150
         TabIndex        =   24
         Top             =   705
         Width           =   2400
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Publisher"
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
         Left            =   4680
         TabIndex        =   22
         Top             =   390
         Width           =   1215
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Author"
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
         Left            =   2520
         TabIndex        =   20
         Top             =   390
         Width           =   975
      End
      Begin VB.OptionButton Opt1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Title"
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
         Left            =   3720
         TabIndex        =   21
         Top             =   390
         Width           =   855
      End
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
         Left            =   1440
         TabIndex        =   19
         Top             =   390
         Width           =   975
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Book Entry Form.frx":0764
         Height          =   4575
         Left            =   120
         TabIndex        =   26
         Top             =   1200
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   8070
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   495
         Left            =   840
         Top             =   3960
         Width           =   1815
         _ExtentX        =   3201
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
         Connect         =   $"Book Entry Form.frx":0779
         OLEDBString     =   $"Book Entry Form.frx":0805
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "select * from book"
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
      Begin VB.Label LblSearch 
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
         Left            =   1635
         TabIndex        =   23
         Top             =   765
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtQty 
      Alignment       =   1  'Right Justify
      DataField       =   "Qty"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   12
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox TxtPrice 
      Alignment       =   1  'Right Justify
      DataField       =   "Price"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox TxtAuthor 
      DataField       =   "Author"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   5
      Top             =   1920
      Width           =   4695
   End
   Begin VB.TextBox TxtTitle 
      DataField       =   "Title"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1440
      Width           =   4695
   End
   Begin VB.TextBox TxtCode 
      DataField       =   "Code"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   1
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox TxtPub 
      DataField       =   "Publisher"
      DataSource      =   "Adodc1"
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2400
      Width           =   4695
   End
   Begin VB.CommandButton CmdExit 
      Caption         =   "E&xit"
      Height          =   630
      Left            =   5100
      Picture         =   "Book Entry Form.frx":0891
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Exit"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   630
      Left            =   3420
      Picture         =   "Book Entry Form.frx":0BCF
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Save Record"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "&Delete"
      Height          =   630
      Left            =   2580
      Picture         =   "Book Entry Form.frx":1239
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Delete Record"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   630
      Left            =   1740
      Picture         =   "Book Entry Form.frx":157B
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Edit Record"
      Top             =   5280
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   630
      Left            =   900
      Picture         =   "Book Entry Form.frx":18BD
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Add Record"
      Top             =   5280
      Width           =   735
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOOK OPERATIONS"
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
      Left            =   5197
      TabIndex        =   25
      Top             =   45
      Width           =   3540
   End
   Begin VB.Shape ShapLabel 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   13935
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Total &Quantity :"
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
      TabIndex        =   11
      Top             =   3900
      Width           =   1320
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "P&rice :"
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
      Top             =   3360
      Width           =   555
   End
   Begin VB.Label LblAuther 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Author :"
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
      Top             =   1980
      Width           =   660
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Title :"
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
      TabIndex        =   2
      Top             =   1500
      Width           =   480
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
      TabIndex        =   0
      Top             =   1020
      Width           =   585
   End
   Begin VB.Label LblPurDt 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Purchase &Date :"
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
      TabIndex        =   8
      Top             =   2940
      Width           =   1425
   End
   Begin VB.Label LblPub 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Publisher :"
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
      TabIndex        =   6
      Top             =   2460
      Width           =   930
   End
End
Attribute VB_Name = "frmBkEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private Sub CmdAdd_Click()
Adodc1.Recordset.AddNew
tenable (True)
cenable (False)
End Sub

Private Sub CmdCancel_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
tenable (False)
cenable (True)
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
If TxtCode.Text = "" Or TxtTitle.Text = "" Or TxtAuthor.Text = "" Or TxtPub.Text = "" Or TxtPrice.Text = "" Or TxtQty.Text = "" Then
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
Text1.Text = DTPicker1.Value
End Sub

Private Sub Form_Load()
tenable (False)
cenable (True)
DataGrid1.Columns(5).Width = 500
DataGrid1.Columns(6).Width = 500
End Sub

Private Sub Opt1_Click(Index As Integer)
TxtSearch.Enabled = True
TxtSearch.SetFocus
i = Index
End Sub

Private Sub Text1_Change()
If Not Text1.Text = "" Then DTPicker1.Value = Text1.Text
End Sub

Private Function tenable(a As Boolean)
TxtCode.Enabled = a
TxtTitle.Enabled = a
TxtAuthor.Enabled = a
TxtPub.Enabled = a
TxtPrice.Enabled = a
TxtQty.Enabled = a
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
DataGrid1.Columns(5).Width = 500
DataGrid1.Columns(6).Width = 500
End Function

Private Sub TxtSearch_Change()
Adodc1.RecordSource = "select * from book where " + Opt1(i).Caption + " like '" + TxtSearch.Text + "%'"
Call refr
Call refr
End Sub
