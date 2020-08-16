VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.MDIForm MDIFrm 
   BackColor       =   &H8000000C&
   Caption         =   "Library"
   ClientHeight    =   10635
   ClientLeft      =   120
   ClientTop       =   765
   ClientWidth     =   20250
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   10260
      Width           =   20250
      _ExtentX        =   35719
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Welcome to library management system "
            TextSave        =   "Welcome to library management system "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5477
            TextSave        =   "14-01-2020"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5477
            TextSave        =   "07:21 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   1
            Enabled         =   0   'False
            Object.Width           =   5477
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   1
            Object.Width           =   5477
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Form.frx":0ECA
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pct1 
      Align           =   1  'Align Top
      BackColor       =   &H00C0E0FF&
      Height          =   11415
      Left            =   0
      ScaleHeight     =   11355
      ScaleWidth      =   20190
      TabIndex        =   1
      Top             =   0
      Width           =   20250
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Current Issue Details"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3375
         Left            =   3405
         TabIndex        =   39
         Top             =   6840
         Width           =   13455
         Begin MSDataGridLib.DataGrid DataGrid3 
            Bindings        =   "Main Form.frx":1DA4
            Height          =   2895
            Left            =   120
            TabIndex        =   43
            Top             =   360
            Width           =   13215
            _ExtentX        =   23310
            _ExtentY        =   5106
            _Version        =   393216
            AllowUpdate     =   0   'False
            DefColWidth     =   120
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
         Begin MSAdodcLib.Adodc Adodc3 
            Height          =   615
            Left            =   4320
            Top             =   960
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   1085
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
            Connect         =   $"Main Form.frx":1DB9
            OLEDBString     =   $"Main Form.frx":1E45
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Select * from Issue"
            Caption         =   "Adodc3"
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
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find Member"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   13560
         TabIndex        =   36
         Top             =   360
         Width           =   6615
         Begin VB.TextBox Txtsearch2 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            Height          =   375
            Left            =   2910
            MaxLength       =   15
            TabIndex        =   54
            Top             =   840
            Width           =   2415
         End
         Begin VB.OptionButton Opt2 
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
            Left            =   4560
            TabIndex        =   53
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton Opt2 
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
            Left            =   2400
            TabIndex        =   52
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton Opt2 
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
            Left            =   3600
            TabIndex        =   51
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Opt2 
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
            Left            =   1320
            TabIndex        =   50
            Top             =   480
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid DataGrid2 
            Bindings        =   "Main Form.frx":1ED1
            Height          =   4695
            Left            =   120
            TabIndex        =   38
            Top             =   1440
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   8281
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
         Begin MSAdodcLib.Adodc Adodc2 
            Height          =   375
            Left            =   2760
            Top             =   4080
            Width           =   1215
            _ExtentX        =   2143
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
            Connect         =   $"Main Form.frx":1EE6
            OLEDBString     =   $"Main Form.frx":1F72
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from Member_Query"
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
            Left            =   1320
            TabIndex        =   55
            Top             =   900
            Width           =   1485
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Find Book"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6255
         Left            =   120
         TabIndex        =   35
         Top             =   360
         Width           =   6615
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
            Left            =   1080
            TabIndex        =   48
            Top             =   480
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
            Left            =   3360
            TabIndex        =   47
            Top             =   480
            Width           =   975
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
            Left            =   2160
            TabIndex        =   46
            Top             =   480
            Width           =   975
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
            Left            =   4320
            TabIndex        =   45
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtSearch 
            BackColor       =   &H00C0E0FF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2790
            TabIndex        =   44
            Top             =   795
            Width           =   2400
         End
         Begin MSDataGridLib.DataGrid DataGrid1 
            Bindings        =   "Main Form.frx":1FFE
            Height          =   4695
            Left            =   120
            TabIndex        =   37
            Top             =   1440
            Width           =   6375
            _ExtentX        =   11245
            _ExtentY        =   8281
            _Version        =   393216
            AllowUpdate     =   0   'False
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
         Begin MSAdodcLib.Adodc Adodc1 
            Height          =   375
            Left            =   600
            Top             =   4200
            Width           =   1575
            _ExtentX        =   2778
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
            Connect         =   $"Main Form.frx":2013
            OLEDBString     =   $"Main Form.frx":209F
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "select * from Book_Query"
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
            Left            =   1275
            TabIndex        =   49
            Top             =   855
            Width           =   1485
         End
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   6270
         Left            =   7065
         TabIndex        =   2
         Top             =   360
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   11060
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         BackColor       =   12640511
         ForeColor       =   32768
         TabCaption(0)   =   "&Issue Book"
         TabPicture(0)   =   "Main Form.frx":212B
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "CmdIssue"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "FremBk1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "FremMbr1"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "&Return Book"
         TabPicture(1)   =   "Main Form.frx":2147
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Label11"
         Tab(1).Control(1)=   "Label12"
         Tab(1).Control(2)=   "Label13"
         Tab(1).Control(3)=   "Label15"
         Tab(1).Control(4)=   "Adodc4"
         Tab(1).Control(5)=   "FremBk2"
         Tab(1).Control(6)=   "FremMbr2"
         Tab(1).Control(7)=   "CmdSubmit"
         Tab(1).ControlCount=   8
         Begin VB.Frame FremMbr1 
            Caption         =   "&Member Detail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1875
            Left            =   120
            TabIndex        =   27
            Top             =   390
            Width           =   5910
            Begin VB.TextBox TxtMbrCode1 
               Height          =   375
               Left            =   1680
               TabIndex        =   59
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox TxtName1 
               Height          =   375
               Left            =   1650
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   28
               TabStop         =   0   'False
               Top             =   1155
               Width           =   4065
            End
            Begin VB.Label LblMbrCode1 
               AutoSize        =   -1  'True
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
               Left            =   135
               TabIndex        =   30
               Top             =   555
               Width           =   585
            End
            Begin VB.Label LblName1 
               AutoSize        =   -1  'True
               Caption         =   "Member Name :"
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
               Left            =   150
               TabIndex        =   29
               Top             =   1185
               Width           =   1440
            End
         End
         Begin VB.Frame FremBk1 
            Caption         =   "&Book Detail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3240
            Left            =   120
            TabIndex        =   19
            Top             =   2400
            Width           =   5910
            Begin VB.TextBox Text1 
               Height          =   375
               Left            =   1410
               Locked          =   -1  'True
               TabIndex        =   34
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox TxtBkCode1 
               Height          =   375
               Left            =   1410
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   315
               Width           =   1695
            End
            Begin VB.TextBox TxtTitle1 
               Height          =   375
               Left            =   1410
               Locked          =   -1  'True
               TabIndex        =   20
               TabStop         =   0   'False
               Top             =   915
               Width           =   4320
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   2280
               TabIndex        =   32
               Top             =   2160
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   114884611
               CurrentDate     =   43836
            End
            Begin MSComCtl2.DTPicker DTPicker3 
               Height          =   375
               Left            =   2280
               TabIndex        =   33
               Top             =   2680
               Width           =   2295
               _ExtentX        =   4048
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   114884611
               CurrentDate     =   43836
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Book Title :"
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
               Left            =   180
               TabIndex        =   26
               Top             =   960
               Width           =   1005
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Stock :"
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
               Left            =   180
               TabIndex        =   25
               Top             =   1530
               Width           =   600
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Last Submit Date :"
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
               Left            =   450
               TabIndex        =   24
               Top             =   2775
               Width           =   1605
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Issue Date :"
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
               Left            =   450
               TabIndex        =   23
               Top             =   2205
               Width           =   1050
            End
            Begin VB.Label LblBkCode1 
               AutoSize        =   -1  'True
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
               Left            =   180
               TabIndex        =   22
               Top             =   390
               Width           =   585
            End
         End
         Begin VB.CommandButton CmdIssue 
            Caption         =   "I&ssue"
            Height          =   435
            Left            =   4950
            TabIndex        =   18
            Top             =   5700
            Width           =   1065
         End
         Begin VB.CommandButton CmdSubmit 
            Caption         =   "&Submit"
            Height          =   435
            Left            =   -70050
            TabIndex        =   17
            Top             =   5700
            Width           =   1065
         End
         Begin VB.Frame FremMbr2 
            Caption         =   "&Member Detail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1875
            Left            =   -74880
            TabIndex        =   12
            Top             =   390
            Width           =   5910
            Begin VB.TextBox TxtMbrCode2 
               Height          =   375
               Left            =   1650
               TabIndex        =   14
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox TxtName2 
               Height          =   375
               Left            =   1650
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   13
               Top             =   1155
               Width           =   4065
            End
            Begin VB.Label LblName2 
               AutoSize        =   -1  'True
               Caption         =   "Member Name :"
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
               Left            =   150
               TabIndex        =   16
               Top             =   1185
               Width           =   1440
            End
            Begin VB.Label LblMbrCode2 
               AutoSize        =   -1  'True
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
               Left            =   135
               TabIndex        =   15
               Top             =   555
               Width           =   585
            End
         End
         Begin VB.Frame FremBk2 
            Caption         =   "&Book Detail"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3240
            Left            =   -74880
            TabIndex        =   3
            Top             =   2400
            Width           =   5910
            Begin VB.TextBox Text6 
               Height          =   375
               Left            =   4200
               TabIndex        =   57
               Text            =   "0"
               Top             =   2700
               Width           =   1575
            End
            Begin VB.TextBox Text4 
               Height          =   375
               Left            =   2040
               Locked          =   -1  'True
               TabIndex        =   42
               Top             =   2040
               Width           =   1670
            End
            Begin VB.TextBox Text3 
               Height          =   375
               Left            =   4080
               Locked          =   -1  'True
               TabIndex        =   41
               Top             =   1440
               Width           =   1670
            End
            Begin VB.TextBox Text2 
               Height          =   375
               Left            =   1410
               Locked          =   -1  'True
               TabIndex        =   40
               Top             =   1440
               Width           =   1095
            End
            Begin MSComCtl2.DTPicker DTPicker1 
               Height          =   375
               Left            =   1440
               TabIndex        =   31
               Top             =   2685
               Width           =   1815
               _ExtentX        =   3201
               _ExtentY        =   661
               _Version        =   393216
               CustomFormat    =   "dd-MM-yyyy"
               Format          =   114884611
               CurrentDate     =   43836
            End
            Begin VB.TextBox TxtBkCode2 
               Height          =   375
               Left            =   1410
               Locked          =   -1  'True
               MaxLength       =   6
               TabIndex        =   5
               Top             =   315
               Width           =   1695
            End
            Begin VB.TextBox TxtTitle2 
               Height          =   375
               Left            =   1410
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   4
               Top             =   850
               Width           =   4320
            End
            Begin VB.Label Label7 
               Caption         =   "Incase for damage"
               Height          =   255
               Left            =   4320
               TabIndex        =   58
               Top             =   2400
               Width           =   1335
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               Caption         =   "Fine :"
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
               TabIndex        =   56
               Top             =   2760
               Width           =   480
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Issued Date :"
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
               Left            =   2820
               TabIndex        =   11
               Top             =   1530
               Width           =   1170
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Last Submit Date :"
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
               Left            =   210
               TabIndex        =   10
               Top             =   2085
               Width           =   1605
            End
            Begin VB.Label Label14 
               AutoSize        =   -1  'True
               Caption         =   "Stock :"
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
               Left            =   180
               TabIndex        =   9
               Top             =   1530
               Width           =   600
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               Caption         =   "Book Title :"
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
               Left            =   180
               TabIndex        =   8
               Top             =   930
               Width           =   1005
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Submit Date :"
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
               Left            =   195
               TabIndex        =   7
               Top             =   2745
               Width           =   1185
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
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
               Left            =   180
               TabIndex        =   6
               Top             =   390
               Width           =   585
            End
         End
         Begin MSAdodcLib.Adodc Adodc4 
            Height          =   375
            Left            =   -72240
            Top             =   4800
            Visible         =   0   'False
            Width           =   1935
            _ExtentX        =   3413
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
            Connect         =   $"Main Form.frx":2163
            OLEDBString     =   $"Main Form.frx":21EF
            OLEDBFile       =   ""
            DataSourceName  =   ""
            OtherAttributes =   ""
            UserName        =   ""
            Password        =   ""
            RecordSource    =   "Select * from Fine"
            Caption         =   "Adodc4"
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
         Begin VB.Label Label15 
            Caption         =   "Label15"
            Height          =   375
            Left            =   -69840
            TabIndex        =   63
            Top             =   5760
            Width           =   255
         End
         Begin VB.Label Label13 
            Caption         =   "Per Day it's calculated automatically"
            Height          =   390
            Left            =   -72240
            TabIndex        =   62
            Top             =   5760
            Width           =   1800
            WordWrap        =   -1  'True
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   -72690
            TabIndex        =   61
            Top             =   5700
            Width           =   390
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Fine amount for late return"
            Height          =   195
            Left            =   -74640
            TabIndex        =   60
            Top             =   5760
            Width           =   1845
         End
      End
   End
   Begin VB.Menu MnuMbr 
      Caption         =   "Mem&ber"
      Begin VB.Menu MnuMbrOpr 
         Caption         =   "&Member Operations"
         Shortcut        =   ^M
      End
      Begin VB.Menu fine 
         Caption         =   "Fine Payment"
         Shortcut        =   ^F
      End
   End
   Begin VB.Menu MnuBk 
      Caption         =   "B&ook"
      Begin VB.Menu MnuBkOpr 
         Caption         =   "Book &Operations"
         Shortcut        =   ^B
      End
   End
   Begin VB.Menu MnuRpt 
      Caption         =   "&Report"
      Begin VB.Menu finerep 
         Caption         =   "&Fine Report"
      End
      Begin VB.Menu mnuIsuRpt 
         Caption         =   "&Issue Report"
      End
      Begin VB.Menu MnuMbrRpt 
         Caption         =   "&Member Report"
      End
      Begin VB.Menu MnuBkRpt 
         Caption         =   "&Book Report"
      End
   End
   Begin VB.Menu set 
      Caption         =   "Settings"
   End
End
Attribute VB_Name = "MDIFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i, j As Integer
Dim fine1, def As Single
Private Sub CmdIssue_Click()
If TxtMbrCode1.Text = "" Or TxtBkCode1.Text = "" Then MsgBox "Please fill all the details", vbExclamation: Exit Sub
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Find ("MCode='" + TxtMbrCode1.Text + "'")
End If
If Not Adodc3.Recordset.EOF Then
    MsgBox "Please Return Your Book First", vbExclamation
    Adodc3.Recordset.MoveFirst
Else
    Adodc3.Recordset.AddNew
    Adodc3.Recordset.Fields("MCode") = TxtMbrCode1.Text
    Adodc3.Recordset.Fields("BCode") = TxtBkCode1.Text
    Adodc3.Recordset.Fields("MName") = TxtName1.Text
    Adodc3.Recordset.Fields("BTitle") = Trim(TxtTitle1.Text)
    Adodc3.Recordset.Fields("Isu_Dt") = DTPicker2.Value
    Adodc3.Recordset.Fields("Lst_Dt") = DTPicker3.Value
    Adodc3.Recordset.Fields("Status") = "Not Returned"
    Adodc3.Recordset.Update
    Adodc3.Refresh
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find ("Code='" + TxtBkCode1.Text + "'")
    Adodc1.Recordset.Update "Qty", Adodc1.Recordset.Fields("Qty") - 1
    Adodc1.Refresh
    Adodc1.Recordset.MoveFirst
    Call refr
    Call refr
End If
TxtMbrCode1.Text = ""
TxtBkCode1.Text = ""
End Sub

Private Sub CmdSubmit_Click()
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Find ("MCode='" + TxtMbrCode2.Text + "'")
If Adodc3.Recordset.EOF Then
    MsgBox "Please Enter Valid Details"
    Adodc3.Recordset.MoveFirst
Else
    Adodc3.Recordset.Update "Status", "Returned"
    Adodc3.Refresh
    Adodc1.Recordset.MoveFirst
    Adodc1.Recordset.Find ("Code='" + TxtBkCode2.Text + "'")
    Adodc1.Recordset.Update "Qty", Adodc1.Recordset.Fields("Qty") + 1
    Adodc1.Refresh
    Adodc1.Recordset.MoveFirst
    Call refr
    Call refr
    def = Val(Label12.Caption)
    dt = DTPicker1.Value - CDate(Text4.Text)
    If (dt) > 0 Then
      fine1 = def * dt
      fine1 = fine1 + Val(Text6.Text)
    Else
      fine1 = fine1 + Val(Text6.Text)
    End If
    If fine1 > 0 Then
    MsgBox "Total Fine is " & fine1, vbInformation
    Adodc4.Refresh
    Adodc4.Recordset.AddNew
    Adodc4.Recordset.Fields(0) = TxtMbrCode2.Text
    Adodc4.Recordset.Fields(1) = TxtBkCode2.Text
    Adodc4.Recordset.Fields(2) = fine1
    Adodc4.Recordset.Fields(3) = DTPicker1.Value
    Adodc4.Recordset.Fields(4) = "Not Paid"
    Adodc4.Recordset.Update
    Adodc4.Refresh
    Adodc2.Recordset.MoveFirst
    Adodc2.Recordset.Find ("Code='" + TxtMbrCode2.Text + "'")
    Adodc2.Recordset.Update "Fine", fine1
    Call refr
    Call refr
    End If
    TxtMbrCode2.Text = ""
End If
End Sub

Private Sub DTPicker2_Change()
DTPicker3.Value = DTPicker2.Value + 7
End Sub

Private Sub fine_Click()
Frmfine.Show vbModal, Me
Call refr
Call refr
End Sub

Private Sub finerep_Click()
frmdate.Label3.Caption = "Fine"
frmdate.Show vbModal, Me
End Sub

Private Sub MDIForm_Load()
DTPicker1.Value = Format(Date, "dd-mm-yyyy")
DTPicker2.Value = DTPicker1.Value
DTPicker3.Value = DTPicker2.Value + 7
Adodc3.RecordSource = "Select * from Issue where status ='Not Returned'"
Call refr
Call refr
Adodc4.RecordSource = "Select * from Settings"
Adodc4.Refresh
Label12.Caption = Adodc4.Recordset.Fields(0)
Label15.Caption = Adodc4.Recordset.Fields(1)
Adodc4.RecordSource = "Select * from Fine"
Adodc4.Refresh
fine1 = 0
fine1 = Format(fine1, "Fixed")
End Sub

Private Sub MnuBkOpr_Click()
frmBkEntry.Show vbModal, Me
End Sub

Private Sub MnuBkRpt_Click()
Dim db As New ADODB.Connection
Dim rs As New ADODB.Recordset
db.Open Adodc1.ConnectionString
rs.Open "Select * from Book", db, adOpenKeyset, adLockOptimistic
Set BKReport1.DataSource = rs
BKReport1.Show
End Sub

Private Sub mnuIsuRpt_Click()
frmdate.Label3.Caption = "Issue"
frmdate.Show vbModal, Me
End Sub

Private Sub MnuMbrOpr_Click()
FrmMember.Show vbModal, Me
End Sub

Private Sub MnuMbrRpt_Click()
FrmMbrDtl.Show vbModal, Me
End Sub

Private Sub Opt2_Click(Index As Integer)
Txtsearch2.Enabled = True
Txtsearch2.SetFocus
j = Index
End Sub
Private Sub Opt1_Click(Index As Integer)
TxtSearch.Enabled = True
TxtSearch.SetFocus
i = Index
End Sub

Private Sub set_Click()
frmset.Show vbModal, Me
End Sub

Private Sub TxtBkCode1_Change()
TxtSearch.Text = ""
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("Code='" + TxtBkCode1.Text + "'")
If Not Adodc1.Recordset.EOF Then
    TxtTitle1.Text = Adodc1.Recordset.Fields("Title")
    Text1.Text = Adodc1.Recordset.Fields("Qty")
Else
    TxtTitle1.Text = ""
    Text1.Text = ""
End If
Adodc1.Recordset.MoveFirst
End Sub

Private Sub TxtBkCode2_Change()
TxtSearch.Text = ""
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("Code='" + TxtBkCode2.Text + "'")
If Not Adodc1.Recordset.EOF Then
    Text2.Text = Adodc1.Recordset.Fields("Qty")
Else
    Text2.Text = ""
End If
Adodc1.Recordset.MoveFirst
End Sub

Private Sub TxtMbrCode1_Change()
Txtsearch2.Text = ""
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Find ("Code='" + TxtMbrCode1.Text + "'")
If Not Adodc2.Recordset.EOF Then
    TxtName1.Text = Adodc2.Recordset.Fields("Name")
Else
    TxtName1.Text = ""
End If
Adodc2.Recordset.MoveFirst
End Sub

Private Sub TxtMbrCode1_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub TxtMbrCode2_Change()
Adodc3.Refresh
If Not Adodc3.Recordset.EOF Then
Adodc3.Recordset.MoveFirst
Adodc3.Recordset.Find ("MCode='" + TxtMbrCode2.Text + "'")
End If
If Not Adodc3.Recordset.EOF Then
    TxtName2.Text = Adodc3.Recordset.Fields("MName")
    TxtBkCode2.Text = Adodc3.Recordset.Fields("BCode")
    TxtTitle2.Text = Adodc3.Recordset.Fields("BTitle")
    Text3.Text = Adodc3.Recordset.Fields("Isu_Dt")
    Text4.Text = Adodc3.Recordset.Fields("Lst_Dt")
Else
    TxtName2.Text = ""
    TxtBkCode2.Text = ""
    TxtTitle2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
End If
If Not Adodc3.Recordset.EOF Then Adodc3.Recordset.MoveFirst
End Sub

Private Sub TxtMbrCode2_KeyPress(KeyAscii As Integer)
If KeyAscii > 96 And KeyAscii < 123 Then KeyAscii = KeyAscii - 32
End Sub

Private Sub TxtSearch_Change()
Adodc1.RecordSource = "select * from Book_Query where " + Opt1(i).Caption + " like '" + TxtSearch.Text + "%'"
Call refr
Call refr
End Sub

Private Sub Txtsearch2_Change()
Adodc2.RecordSource = "select * from Member_Query where " + Opt2(j).Caption + " like '" + Txtsearch2.Text + "%'"
Call refr
Call refr
End Sub
Private Function refr()
Adodc1.Refresh
Adodc2.Refresh
Adodc3.Refresh
DataGrid1.Refresh
DataGrid2.Refresh
DataGrid3.Refresh
DataGrid1.Columns(4).Width = 500
DataGrid2.Columns(4).Width = 500
End Function
