VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmBookIsu 
   Caption         =   "Issue/Submit Book/CD"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Book_Isu_Form.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FremBook 
      Caption         =   "Select Book"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   6315
      TabIndex        =   59
      Top             =   3585
      Width           =   5445
      Begin MSFlexGridLib.MSFlexGrid MsfgBook 
         Height          =   2880
         Left            =   75
         TabIndex        =   60
         Top             =   315
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   5080
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame FremMbr 
      Caption         =   "Select Member"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   6315
      TabIndex        =   57
      Top             =   90
      Width           =   5445
      Begin MSFlexGridLib.MSFlexGrid MsfgMbr 
         Height          =   2865
         Left            =   75
         TabIndex        =   58
         Top             =   315
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   5054
         _Version        =   393216
         AllowUserResizing=   3
      End
   End
   Begin VB.Frame Frame5 
      Height          =   660
      Left            =   30
      TabIndex        =   0
      Top             =   120
      Width           =   6120
      Begin VB.ComboBox CmbType 
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
         ItemData        =   "Book_Isu_Form.frx":030A
         Left            =   780
         List            =   "Book_Isu_Form.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   195
         Width           =   975
      End
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
         ItemData        =   "Book_Isu_Form.frx":0322
         Left            =   4965
         List            =   "Book_Isu_Form.frx":0341
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   945
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
         ItemData        =   "Book_Isu_Form.frx":0375
         Left            =   2910
         List            =   "Book_Isu_Form.frx":0391
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   1080
      End
      Begin VB.Label LblType 
         AutoSize        =   -1  'True
         Caption         =   "&Type :"
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
         TabIndex        =   1
         Top             =   255
         Width           =   570
      End
      Begin VB.Label LblClassYear 
         AutoSize        =   -1  'True
         Caption         =   "&Year :"
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
         Left            =   4410
         TabIndex        =   5
         Top             =   255
         Width           =   525
      End
      Begin VB.Label LblClass 
         AutoSize        =   -1  'True
         Caption         =   "C&lass :"
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
         Left            =   2295
         TabIndex        =   3
         Top             =   255
         Width           =   600
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6390
      Left            =   0
      TabIndex        =   7
      Top             =   990
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   11271
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Issue Book"
      TabPicture(0)   =   "Book_Isu_Form.frx":03C4
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FremMbr1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FremBk1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdIssue"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&Submit Book"
      TabPicture(1)   =   "Book_Isu_Form.frx":03E0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "CmdSubmit"
      Tab(1).Control(1)=   "FremMbr2"
      Tab(1).Control(2)=   "FremBk2"
      Tab(1).ControlCount=   3
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
         TabIndex        =   38
         Top             =   2325
         Width           =   5910
         Begin VB.TextBox TxtTitle2 
            Height          =   375
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   42
            Top             =   915
            Width           =   4320
         End
         Begin VB.TextBox TxtBkCode2 
            Height          =   375
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   40
            Top             =   315
            Width           =   1095
         End
         Begin VB.TextBox TxtLastDt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4635
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   2130
            Width           =   1095
         End
         Begin VB.TextBox TxtIsuedDt 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   48
            Top             =   2145
            Width           =   1095
         End
         Begin VB.ComboBox CmbYear3 
            Height          =   315
            Left            =   3285
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   2730
            Width           =   855
         End
         Begin VB.ComboBox CmbMonth3 
            Height          =   315
            Left            =   2565
            Style           =   2  'Dropdown List
            TabIndex        =   53
            Top             =   2730
            Width           =   735
         End
         Begin VB.TextBox TxtAvlStock2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4635
            Locked          =   -1  'True
            TabIndex        =   46
            Top             =   1470
            Width           =   1095
         End
         Begin VB.TextBox TxtTotStock2 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   44
            Top             =   1470
            Width           =   1095
         End
         Begin VB.ComboBox CmbDay3 
            Height          =   315
            Left            =   1845
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   2730
            Width           =   735
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
            TabIndex        =   39
            Top             =   390
            Width           =   585
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
            Left            =   555
            TabIndex        =   51
            Top             =   2745
            Width           =   1185
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
            TabIndex        =   41
            Top             =   810
            Width           =   1005
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "Total Stock :"
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
            TabIndex        =   43
            Top             =   1530
            Width           =   1110
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            Caption         =   "Available Stock :"
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
            Left            =   2970
            TabIndex        =   45
            Top             =   1530
            Width           =   1500
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
            Left            =   2970
            TabIndex        =   49
            Top             =   2190
            Width           =   1605
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
            Left            =   180
            TabIndex        =   47
            Top             =   2205
            Width           =   1170
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "DD-MMM-YYYY"
            Height          =   195
            Left            =   4185
            TabIndex        =   55
            Top             =   2790
            Width           =   1155
         End
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
         TabIndex        =   33
         Top             =   390
         Width           =   5910
         Begin VB.TextBox TxtName2 
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   37
            Top             =   1155
            Width           =   4065
         End
         Begin VB.TextBox TxtMbrCode2 
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   35
            Top             =   480
            Width           =   1095
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
            TabIndex        =   34
            Top             =   555
            Width           =   585
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
            TabIndex        =   36
            Top             =   1185
            Width           =   1440
         End
      End
      Begin VB.CommandButton CmdSubmit 
         Caption         =   "Submit"
         Height          =   435
         Left            =   -70050
         TabIndex        =   56
         Top             =   5760
         Width           =   1065
      End
      Begin VB.CommandButton CmdIssue 
         Caption         =   "&Issue"
         Height          =   435
         Left            =   4950
         TabIndex        =   32
         Top             =   5760
         Width           =   1065
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
         TabIndex        =   13
         Top             =   2325
         Width           =   5910
         Begin VB.TextBox TxtTitle1 
            Height          =   375
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   915
            Width           =   4320
         End
         Begin VB.TextBox TxtBkCode1 
            Height          =   375
            Left            =   1410
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   315
            Width           =   1095
         End
         Begin VB.ComboBox CmbYear2 
            Height          =   315
            Left            =   3540
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   2730
            Width           =   855
         End
         Begin VB.ComboBox CmbMonth2 
            Height          =   315
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   2730
            Width           =   735
         End
         Begin VB.ComboBox CmbDay2 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   2730
            Width           =   735
         End
         Begin VB.ComboBox CmbYear1 
            Height          =   315
            Left            =   3540
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   2175
            Width           =   855
         End
         Begin VB.ComboBox CmbMonth1 
            Height          =   315
            Left            =   2820
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   2175
            Width           =   735
         End
         Begin VB.ComboBox CmbDay1 
            Height          =   315
            Left            =   2100
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2175
            Width           =   735
         End
         Begin VB.TextBox TxtTotStock1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1410
            Locked          =   -1  'True
            TabIndex        =   19
            TabStop         =   0   'False
            Top             =   1470
            Width           =   1095
         End
         Begin VB.TextBox TxtAvlStock1 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4635
            Locked          =   -1  'True
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   1470
            Width           =   1095
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
            TabIndex        =   14
            Top             =   390
            Width           =   585
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            Caption         =   "DD-MM-YYYY"
            Height          =   195
            Left            =   4455
            TabIndex        =   31
            Top             =   2790
            Width           =   1020
         End
         Begin VB.Label LblDtFrmt 
            AutoSize        =   -1  'True
            Caption         =   "DD-MM-YYYY"
            Height          =   195
            Left            =   4440
            TabIndex        =   26
            Top             =   2235
            Width           =   1020
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
            TabIndex        =   22
            Top             =   2205
            Width           =   1050
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
            TabIndex        =   27
            Top             =   2775
            Width           =   1605
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Available Stock :"
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
            Left            =   3075
            TabIndex        =   20
            Top             =   1530
            Width           =   1500
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Total Stock :"
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
            TabIndex        =   18
            Top             =   1530
            Width           =   1110
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
            TabIndex        =   16
            Top             =   960
            Width           =   1005
         End
      End
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
         TabIndex        =   8
         Top             =   390
         Width           =   5910
         Begin VB.TextBox TxtName1 
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   1155
            Width           =   4065
         End
         Begin VB.TextBox TxtMbrCode1 
            Height          =   375
            Left            =   1650
            Locked          =   -1  'True
            MaxLength       =   6
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   480
            Width           =   1095
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
            TabIndex        =   11
            Top             =   1185
            Width           =   1440
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
            TabIndex        =   9
            Top             =   555
            Width           =   585
         End
      End
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   390
      Left            =   10695
      TabIndex        =   61
      Top             =   6990
      Width           =   1065
   End
End
Attribute VB_Name = "FrmBookIsu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_mbr As New ADODB.Recordset
Dim rs_bk As New ADODB.Recordset
Dim rs_isu As New ADODB.Recordset
Dim rs_fine As New ADODB.Recordset

Dim Qry, dt, ldt As String, fine As Integer


Private Sub CmbClass_Click()
    Call fillYear(Me) 'SELECT YEAR
    CmbClassYear.Text = CmbClassYear.List(0)
End Sub


Private Sub CmbClassYear_Click()
    
    'RETRIVE RECORDS
    Call SSTab1_Click(SSTab1.Tab)
    
    'CLEAR TEXT BOXES
    Call clearText
    
    CmbDay1.Text = Day(Date)
    CmbMonth1.Text = Month(Date)
    CmbYear1.Text = Year(Date)

    CmbDay2.Text = Day(DateAdd("d", 7, CmbDay1.Text & "-" & CmbMonth1.Text & "-" & CmbYear1.Text))
    CmbMonth2.Text = Month(DateAdd("d", 7, Date))
    CmbYear2.Text = Year(DateAdd("d", 7, Date))

    CmbDay3.Text = Day(Date)
    CmbMonth3.Text = Month(Date)
    CmbYear3.Text = Year(Date)
        
End Sub

Private Sub CmbDay1_Click()
    'FILL LAST SUBMIT YEAR COMBO
    CmbYear2.Clear
    For i = Val(CmbYear1.Text) To 2051
        CmbYear2.AddItem i
    Next
    CmbYear2.Text = CmbYear2.List(0)
    
    If CmbDay1.Text <> "" And CmbMonth1.Text <> "" And CmbYear1.Text <> "" Then
        CmbYear2.Text = Year(DateAdd("d", 7, CmbDay1.Text & "-" & CmbMonth1.Text & "-" & CmbYear1.Text))
        CmbMonth2.Text = Month(DateAdd("d", 7, CmbDay1.Text & "-" & CmbMonth1.Text & "-" & CmbYear1.Text))
        CmbDay2.Text = Day(DateAdd("d", 7, CmbDay1.Text & "-" & CmbMonth1.Text & "-" & CmbYear1.Text))
    End If
End Sub

Private Sub CmbMonth1_Click()
    Dim i As Integer
    CmbDay1.Clear
    For i = 1 To daysOfMonth(Val(CmbMonth1.Text), Val(CmbYear1.Text))
        CmbDay1.AddItem i
    Next i
    CmbDay1.Text = Day(Date)

End Sub

Private Sub CmbMonth2_Click()
    Dim i As Integer, dt As String
    CmbDay2.Clear
    dt = CmbDay1.Text & "/" & CmbMonth1.Text & "/" & CmbYear1.Text
    For i = 1 To DateDiff("d", DateAdd("d", 7, dt), DateAdd("m", 1, DateAdd("d", 7, dt)))
        CmbDay2.AddItem i
    Next
    CmbDay2.Text = Day(DateAdd("d", 7, dt))
End Sub

Private Sub CmbMonth3_Click()
    Dim i As Integer
    CmbDay3.Clear
    For i = 1 To daysOfMonth(Val(CmbMonth3.Text), Val(CmbYear3.Text))
        CmbDay3.AddItem i
    Next i
    CmbDay3.Text = Day(Date)
End Sub

Private Sub CmbType_Click()
    
    Call SSTab1_Click(SSTab1.Tab)

End Sub

Private Sub CmbYear1_Click()
    Dim i As Integer
    CmbDay1.Clear
    For i = 1 To daysOfMonth(Val(CmbMonth1.Text), Val(CmbYear1.Text))
        CmbDay1.AddItem i
    Next i
    CmbDay1.Text = Day(Date)
    
    
    'FILL LAST SUBMIT YEAR COMBO
    CmbYear2.Clear
    For i = Val(CmbYear1.Text) To 2051
        CmbYear2.AddItem i
    Next
    CmbYear2.Text = CmbYear2.List(0)

End Sub

Private Sub CmbYear3_Click()
    Dim i As Integer
    CmbDay3.Clear
    For i = 1 To daysOfMonth(Val(CmbMonth3.Text), Val(CmbYear3.Text))
        CmbDay3.AddItem i
    Next i
    CmbDay3.Text = Day(Date)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdIssue_Click()
    Dim rs_tmp As New Recordset
    Dim tmpYr As String
        
    'CHECK FOR BLANCK DATA
    If TxtMbrCode1.Text = "" Then
        MsgBox "Enter Member Number.", vbInformation, "Issue Book/CD"
        TxtMbrCode1.SetFocus
        Exit Sub
    End If
    
    If TxtBkCode1.Text = "" Then
        MsgBox "Enter Book Number.", vbInformation, "Issue Book/CD"
        TxtBkCode1.SetFocus
        Exit Sub
    End If
        
    'WHEN NO MEMBER EXIST
    If rs_mbr.RecordCount = 0 Then
        MsgBox "No Member Exist.", vbInformation, "Issue Book/CD"
        Exit Sub
    End If
    
    'WHEN NO BOOK/CD EXIST
    If rs_bk.RecordCount = 0 Then
        MsgBox "No Book Exist.", vbInformation, "Issue Book/CD"
        Exit Sub
    End If
    
    rs_mbr.MoveFirst
    rs_mbr.Find "Code='" & TxtMbrCode1 & "'"
    If rs_mbr.EOF Then
        MsgBox "Member does not exist", vbInformation, "Issue Book/CD"
        TxtMbrCode1.SetFocus
        Exit Sub
    End If
        
    rs_bk.MoveFirst
    rs_bk.Find "Code='" & TxtBkCode1 & "'"
    If rs_bk.EOF Then
        MsgBox "Book does not exist", vbInformation, "Issue Book/CD"
        TxtBkCode1.SetFocus
        Exit Sub
    End If
    
    If TxtAvlStock1 = 0 Then
        MsgBox "No book in stack.", vbInformation, "Issue Book/CD"
        Exit Sub
    End If
    
    'FIND RECORD IN ISSUE_TABLE, IF ANY BOOK ISSUED TO THIS PERSON
    Set rs_tmp = New Recordset
    rs_tmp.Open "SELECT * FROM Issue_Mast WHERE [Mbr_No]='" & TxtMbrCode1 & "' AND [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "'", conn, adOpenStatic
    
    If rs_tmp.RecordCount > 0 Then
        If MsgBox(rs_tmp.RecordCount & " books are already issued." & vbCrLf & _
                "Do you want to issue book ?", vbQuestion + vbYesNo, "Issue Book/CD") = vbNo Then
            
            Exit Sub
        End If
        
        rs_tmp.MoveFirst
        While Not rs_tmp.EOF
            If rs_tmp.Fields(3) = TxtBkCode1 Then
                MsgBox "You can not issued same book to same person.", vbInformation, "Issue Book/CD"
                Exit Sub
            End If
            rs_tmp.MoveNext
        Wend
    End If
    
    
    dt = CmbDay1.Text & "-" & CmbMonth1.Text & "-" & CmbYear1.Text
    ldt = CmbDay2.Text & "-" & CmbMonth2.Text & "-" & CmbYear2.Text
    
    Qry = "insert into Issue_Mast values ('" & TxtMbrCode1 & "','" & CmbClass.Text & "','" & _
        CmbClassYear.Text & "','" & TxtBkCode1 & "','" & dt & "','" & ldt & "')"
        
    conn.Execute Qry
    
    rs_bk.MoveFirst
    rs_bk.Find "Code='" & TxtBkCode1.Text & "'"
    rs_bk.Fields(7) = rs_bk.Fields(7) + 1
    rs_bk.Update
    
    MsgBox "Book or CD is issued.", vbInformation, "Issue Book/CD"
    tmpYr = CmbClassYear.Text
    
    Call CmbType_Click 'TO RETRIVE UPDTED RECORD
    Call CmbClass_Click 'TO RETRIVE MEMBERS
    CmbClassYear.Text = tmpYr 'TO RETRIVE MEMBERS OF PREVIOUS TYPE
End Sub

Private Sub CmdSubmit_Click()
    
    'CHECK FOR BLANCK DATA
    If Trim(TxtMbrCode2) = "" Then
        MsgBox "Enter member code.", vbInformation, "Submit Book/CD"
        TxtMbrCode2.SetFocus
        Exit Sub
    End If
    If Trim(TxtBkCode2) = "" Then
        MsgBox "Enter Book code.", vbInformation, "Submit Book/CD"
        TxtBkCode2.SetFocus
        Exit Sub
    End If
    
    'WHEN NO MEMBER EXIST
    If rs_mbr.RecordCount = 0 Then
        MsgBox "No Member Exist.", vbInformation, "Issue Book/CD"
        Exit Sub
    End If
    
    'WHEN NO BOOK/CD EXIST
    If rs_bk.RecordCount = 0 Then
        MsgBox "No Book Exist.", vbInformation, "Issue Book/CD"
        Exit Sub
    End If
    
    'CHECK IF MEMBER EXIST OR NOT
    rs_mbr.MoveFirst
    rs_mbr.Find "Code='" & TxtMbrCode2 & "'"
    If rs_mbr.EOF Then
        MsgBox "Member does not exist", vbInformation, "Issue Book/CD"
        TxtMbrCode1.SetFocus
        Exit Sub
    End If
        
    'CHECK IF BOOK EXIST OR NOT
    rs_bk.MoveFirst
    rs_bk.Find "Code='" & TxtBkCode2 & "'"
    If rs_bk.EOF Then
        MsgBox "Book does not exist", vbInformation, "Issue Book/CD"
        TxtBkCode1.SetFocus
        Exit Sub
    End If
    
    'CALCULATE FOR FINE
    dt = CmbDay3.Text & "/" & CmbMonth3.Text & "/" & CmbYear3.Text
    If DateCmp(TxtLastDt, dt) = 2 Then
        If MsgBox("Fine is applicable. Do you want to collect it or not ?", vbQuestion + vbYesNo, "Member Fine") = vbYes Then
            
            fine = Val(InputBox("Enter Amount of fine.", "Member Fine", DateDiff("d", TxtLastDt, dt) * 10))
                
            Qry = "UPDATE Mbr_Mast SET [Fine]=[Fine]+" & fine & " WHERE [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & _
                "' AND [Code]='" & TxtMbrCode2 & "'"
            
            conn.Execute Qry
                    
        End If
    End If
    
    dt = CmbDay3.Text & "/" & CmbMonth3.Text & "/" & CmbYear3.Text
    'UPDATE RECORD IN ISSUE_MAST TABLE
    Qry = "DELETE FROM Issue_Mast WHERE [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "' AND [Mbr_No]='" & TxtMbrCode2 & "' AND [Bk_No]='" & TxtBkCode2 & "'"
    
    conn.Execute Qry
    
    'UPDATE BOOK_MAST TABLE
    rs_bk.MoveFirst
    rs_bk.Find "Code='" & TxtBkCode2 & "'"
    
    rs_bk.Fields(7) = rs_bk.Fields(7) - 1
    rs_bk.Update
    
    MsgBox "Book/CD submited successfully.", vbInformation, "Submit Book/CD"
    tmpYr = CmbClassYear.Text
    
    Call CmbClass_Click 'TO RETRIVE MEMBERS
    CmbClassYear.Text = tmpYr 'TO RETRIVE MEMBERS OF PREVIOUS TYPE
End Sub

Private Sub Form_Load()
    Dim i As Integer

    Call fillDate
    
    CmbType.Text = CmbType.List(0)
    CmbClass.Text = CmbClass.List(0)
    
    'SET HEADING OF FLEX GRIDS
    Me.MsfgMbr.FormatString = "No. |Code     |Surname         |Member Name         |Father Name          |Join Date    |City           " & _
                                "|Contect No.    |Gender"
                                
    Me.MsfgBook.FormatString = "No. |Code     |Title                    |Auther                  |Publisher                " & _
                                "|Purchase Date |Price   |Quantity|Issued Book/CD(s)"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Forms.Count = 2 Then
        MDIFrm.Pct1.Visible = True
    End If
End Sub

Private Sub MsfgBook_Click()
    If SSTab1.Tab = 0 Then
        TxtBkCode1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 1)
        TxtTitle1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 2)
        TxtTotStock1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7)
        TxtAvlStock1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7) - MsfgBook.TextMatrix(MsfgBook.Row, 8)
    Else
        TxtBkCode2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 1)
        TxtTitle2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 2)
        TxtTotStock2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7)
        TxtAvlStock2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7) - MsfgBook.TextMatrix(MsfgBook.Row, 8)
        
        If rs_isu.RecordCount > 0 Then
            rs_isu.MoveFirst
            rs_isu.Find "Bk_No='" & TxtBkCode2 & "'"
        
            TxtIsuedDt.Text = Format(rs_isu.Fields(4), "dd/mm/yyyy")
            TxtLastDt.Text = Format(rs_isu.Fields(5), "dd/mm/yyyy")
        End If
        
    End If
End Sub

Private Sub MsfgBook_EnterCell()

    If SSTab1.Tab = 0 Then
        TxtBkCode1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 1)
        TxtTitle1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 2)
        TxtTotStock1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7)
        TxtAvlStock1.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7) - MsfgBook.TextMatrix(MsfgBook.Row, 8)
    Else
        TxtBkCode2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 1)
        TxtTitle2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 2)
        TxtTotStock2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7)
        TxtAvlStock2.Text = MsfgBook.TextMatrix(MsfgBook.Row, 7) - MsfgBook.TextMatrix(MsfgBook.Row, 8)
        
        If rs_isu.RecordCount > 0 Then
            rs_isu.MoveFirst
            rs_isu.Find "Bk_No='" & TxtBkCode2 & "'"
        
            TxtIsuedDt.Text = Format(rs_isu.Fields(4), "dd/mm/yyyy")
            TxtLastDt.Text = Format(rs_isu.Fields(5), "dd/mm/yyyy")
        End If
        
    End If
End Sub

Private Sub MsfgMbr_Click()
    Dim i As Integer
    
    'CLEAR TEXT BOXES
    Call clearText

    If SSTab1.Tab = 0 Then
        TxtMbrCode1.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 1)
        TxtName1.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 2) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 3) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 4)
    Else
        
        TxtMbrCode2.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 1)
        TxtName2.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 2) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 3) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 4)
        
        'RETRIVE ISSUED BOOKS OF PERTICULAR STUDENT
        Set rs_isu = New Recordset
        rs_isu.Open "SELECT * FROM Issue_Mast where [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "' AND [Mbr_No]='" & TxtMbrCode2 & "'", conn, adOpenKeyset
        
        MsfgBook.Cols = rs_bk.Fields.Count + 1
        MsfgBook.Rows = rs_isu.RecordCount + 1
        
        If rs_isu.RecordCount = 0 Then
            FremBook.Enabled = False
        Else
            FremBook.Enabled = True
        End If
        
        For i = 1 To rs_isu.RecordCount
            rs_bk.MoveFirst
            rs_bk.Find "Code='" & rs_isu.Fields(3) & "'"
                                
            MsfgBook.TextMatrix(i, 0) = i
            MsfgBook.TextMatrix(i, 1) = rs_bk.Fields(0)
            MsfgBook.TextMatrix(i, 2) = rs_bk.Fields(1)
            MsfgBook.TextMatrix(i, 3) = rs_bk.Fields(2)
            MsfgBook.TextMatrix(i, 4) = rs_bk.Fields(3)
            MsfgBook.TextMatrix(i, 5) = rs_bk.Fields(4)
            MsfgBook.TextMatrix(i, 6) = rs_bk.Fields(5)
            MsfgBook.TextMatrix(i, 7) = rs_bk.Fields(6)
            MsfgBook.TextMatrix(i, 8) = rs_bk.Fields(7)
            
            rs_isu.MoveNext
        Next i
    
    End If
    
End Sub

Private Sub MsfgMbr_EnterCell()
    Dim i As Integer

    If SSTab1.Tab = 0 Then
        TxtMbrCode1.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 1)
        TxtName1.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 2) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 3) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 4)
    Else
        TxtMbrCode2.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 1)
        TxtName2.Text = MsfgMbr.TextMatrix(MsfgMbr.Row, 2) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 3) & " " & MsfgMbr.TextMatrix(MsfgMbr.Row, 4)
        
        'RETRIVE ISSUED BOOKS OF PERTICULAR STUDENT
        Set rs_isu = New Recordset
        rs_isu.Open "SELECT * FROM Issue_Mast where [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "' AND Mbr_No='" & TxtMbrCode2 & "'", conn, adOpenKeyset
        
        MsfgBook.Cols = rs_bk.Fields.Count + 1
        MsfgBook.Rows = rs_isu.RecordCount + 1
        
        If rs_isu.RecordCount = 0 Then
            FremBook.Enabled = False
        Else
            FremBook.Enabled = True
        End If
        
        For i = 1 To rs_isu.RecordCount
            rs_bk.MoveFirst
            rs_bk.Find "Code='" & rs_isu.Fields(3) & "'"
                                
            MsfgBook.TextMatrix(i, 0) = i
            MsfgBook.TextMatrix(i, 1) = rs_bk.Fields(0)
            MsfgBook.TextMatrix(i, 2) = rs_bk.Fields(1)
            MsfgBook.TextMatrix(i, 3) = rs_bk.Fields(2)
            MsfgBook.TextMatrix(i, 4) = rs_bk.Fields(3)
            MsfgBook.TextMatrix(i, 5) = rs_bk.Fields(4)
            MsfgBook.TextMatrix(i, 6) = rs_bk.Fields(5)
            MsfgBook.TextMatrix(i, 7) = rs_bk.Fields(6)
            MsfgBook.TextMatrix(i, 8) = rs_bk.Fields(7)
            
            rs_isu.MoveNext
        Next i
    
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Dim i As Integer
    
    Set rs_bk = New Recordset
    Set rs_mbr = New Recordset
    Set rs_isu = New Recordset
    
    'CLEAR TEXT BOXES
    Call clearText
    
    
    'RETRIVE MEMBER
    rs_mbr.Open "select Code,surname,member,father,Join_Dt,City,Cnt_No,Gender from Mbr_Mast where crs='" & CmbClass.Text & _
            "' and Yer='" & CmbClassYear.Text & "' ORDER BY Code", conn, adOpenStatic, adLockPessimistic
            
    If rs_mbr.RecordCount = 0 Then
        FremMbr.Enabled = False
    Else
        FremMbr.Enabled = True
    End If
            
    Call fillFlex(MsfgMbr, rs_mbr) 'FILL FLEX GRID
    
    
    If SSTab1.Tab = 0 Then
        Me.Caption = "Issue Book/CD"
        
        FremMbr1.Enabled = True
        FremBk1.Enabled = True
        CmdIssue.Enabled = True
        CmdIssue.Default = True
    
        FremMbr2.Enabled = False
        FremBk2.Enabled = False
        CmdSubmit.Enabled = False
        
        'RETRIVE BOOKS
        If CmbType.Text = "BOOK" Then
            rs_bk.Open "select Code,Title,Author,Publisher,Pur_Dt,Price,Qty,IsudBk from Book_Mast where code like('B%')", conn, adOpenStatic, adLockPessimistic
        Else
            rs_bk.Open "select Code,Title,Author,Publisher,Pur_Dt,Price,Qty,IsudBk from Book_Mast where code like('C%')", conn, adOpenStatic, adLockPessimistic
        End If
    
        If rs_bk.RecordCount = 0 Then
            FremBook.Enabled = False
        Else
            FremBook.Enabled = True
        End If
        
        Call fillFlex(MsfgBook, rs_bk) 'FILL FLEX GRID
      
    Else
        Me.Caption = "Submit Book/CD"
        
        FremMbr2.Enabled = True
        FremBk2.Enabled = True
        CmdSubmit.Enabled = True
        CmdSubmit.Default = True
        
        FremMbr1.Enabled = False
        FremBk1.Enabled = False
        CmdIssue.Enabled = False
        
        'OPEN BOOK RECORDS
        rs_bk.Open "select Code,Title,Author,Publisher,Pur_Dt,Price,Qty,IsudBk from Book_Mast", conn, adOpenStatic, adLockPessimistic
        
        
        'WHEN NO BOOK IS ENTERED THEN NO BOOK IS DISPLAYED IN FLEX GRID
        If rs_bk.RecordCount = 0 Then
            MsfgBook.Cols = rs_bk.Fields.Count + 1
            MsfgBook.Rows = rs_bk.RecordCount + 1
            FremBook.Enabled = False
            Exit Sub
        Else
            FremBook.Enabled = True
        End If
        
        'RETRIVE ISSUED BOOKS/CDS
        rs_isu.Open "SELECT * FROM Issue_Mast WHERE [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "'", conn, adOpenStatic
        
        'WHEN NO BOOK IS ISSSUED THEN
        If rs_isu.RecordCount = 0 Then
            MsfgBook.Cols = rs_bk.Fields.Count + 1
            MsfgBook.Rows = rs_isu.RecordCount + 1
            FremBook.Enabled = False
            Exit Sub
        Else
            FremBook.Enabled = True
        End If
        
        
        'RETRIVE DATA OF ISSUED BOOKS/CD FROM BOOK_MASTER
        MsfgBook.Cols = rs_bk.Fields.Count + 1
        MsfgBook.Rows = rs_isu.RecordCount + 1
        
        rs_isu.MoveFirst
        For i = 1 To rs_isu.RecordCount
            rs_bk.MoveFirst
            rs_bk.Find "Code='" & rs_isu.Fields(3) & "'"
            
            MsfgBook.TextMatrix(i, 0) = i
            MsfgBook.TextMatrix(i, 1) = rs_bk.Fields(0)
            MsfgBook.TextMatrix(i, 2) = rs_bk.Fields(1)
            MsfgBook.TextMatrix(i, 3) = rs_bk.Fields(2)
            MsfgBook.TextMatrix(i, 4) = rs_bk.Fields(3)
            MsfgBook.TextMatrix(i, 5) = rs_bk.Fields(4)
            MsfgBook.TextMatrix(i, 6) = rs_bk.Fields(5)
            MsfgBook.TextMatrix(i, 7) = rs_bk.Fields(6)
            MsfgBook.TextMatrix(i, 8) = rs_bk.Fields(7)
            
            rs_isu.MoveNext
        Next i
                
    End If
End Sub

Private Sub TxtBkCode1_KeyPress(KeyAscii As Integer)
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtBkCode2_KeyPress(KeyAscii As Integer)
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtMbrCode1_KeyPress(KeyAscii As Integer)
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtMbrCode2_KeyPress(KeyAscii As Integer)
    KeyAscii = Book.upper(KeyAscii)
End Sub

'============================================================
'         P R O C E D U R E S
'============================================================

'CLEAR TEXT BOXES
Private Sub clearText()
    Dim cnt As Control
    With Me
        For Each cnt In .Controls
            If TypeOf cnt Is TextBox Then
                cnt.Text = ""
            End If
        Next cnt
    End With
End Sub

'============================================================
'FILL DATE COMBO
Private Sub fillDate()
    Dim i As Integer
    With Me
        CmbYear1.Clear
        For i = 1950 To 2050
            CmbYear1.AddItem i
            CmbYear2.AddItem i
            CmbYear3.AddItem i
        Next
        
        CmbMonth1.Clear
        For i = 1 To 12
            CmbMonth1.AddItem i
            CmbMonth2.AddItem i
            CmbMonth3.AddItem i
        Next
        
        CmbDay1.Clear
        For i = 1 To DateDiff("d", Date, DateAdd("m", 1, Date))
            CmbDay1.AddItem i
            CmbDay3.AddItem i
        Next
        
        CmbDay2.Clear
        For i = 1 To DateDiff("d", DateAdd("m", 1, Date), DateAdd("m", 2, Date))
            CmbDay2.AddItem i
        Next
        
    End With
End Sub
