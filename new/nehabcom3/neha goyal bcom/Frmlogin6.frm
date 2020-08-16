VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form5"
   ClientHeight    =   10935
   ClientLeft      =   195
   ClientTop       =   525
   ClientWidth     =   11025
   LinkTopic       =   "Form5"
   MDIChild        =   -1  'True
   ScaleHeight     =   10935
   ScaleWidth      =   11025
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   450
      Left            =   1560
      Top             =   9120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   794
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
      BackColor       =   12640511
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Frmlogin6.frx":0000
      OLEDBString     =   $"Frmlogin6.frx":0091
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "SR"
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
   Begin VB.CommandButton Command6 
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   20
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "UPDATE"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   19
      Top             =   9600
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "CANCEL"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   18
      Top             =   9600
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      DataField       =   "RNO"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   17
      Text            =   "Text7"
      Top             =   7920
      Width           =   5175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "DELETE"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   16
      Top             =   8640
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Caption         =   "EDIT"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4560
      TabIndex        =   15
      Top             =   8640
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "ADD NEW"
      BeginProperty Font 
         Name            =   "Bodoni MT Condensed"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   14
      Top             =   8640
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      DataField       =   "SECTION"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   13
      Text            =   "Text6"
      Top             =   6840
      Width           =   5175
   End
   Begin VB.TextBox Text5 
      DataField       =   "CLASS"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   12
      Text            =   "Text5"
      Top             =   5760
      Width           =   5175
   End
   Begin VB.TextBox Text4 
      DataField       =   "LNAME"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   11
      Text            =   "Text4"
      Top             =   4680
      Width           =   5175
   End
   Begin VB.TextBox Text3 
      DataField       =   "MNAME"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   10
      Text            =   "Text3"
      Top             =   3600
      Width           =   5175
   End
   Begin VB.TextBox Text2 
      DataField       =   "FNAME"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   2520
      Width           =   5175
   End
   Begin VB.TextBox Text1 
      DataField       =   "SID"
      DataSource      =   "Adodc1"
      Height          =   405
      Left            =   4920
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1440
      Width           =   5175
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ROLL NO."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   7800
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SECTION"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   6
      Top             =   6720
      Width           =   2415
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CLASS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   5640
      Width           =   2415
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LAST NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   4
      Top             =   4560
      Width           =   2415
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MIDDLE NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   3
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "FIRST NAME"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2400
      Width           =   2415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STUDENT ID"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "STUDENT RECORD"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2160
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Refresh
Adodc1.Recordset.AddNew
txt (True)
cmd (False)
End Sub

Private Sub Command2_Click()
txt (True)
cmd (False)
Command4.Enabled = False
End Sub

Private Sub Command3_Click()
If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then
    Adodc1.Recordset.Delete
    If Adodc1.Recordset.BOF Then
        Adodc1.Recordset.MoveFirst
    Else
        Adodc1.Recordset.MovePrevious
    End If
End If
End Sub

Private Sub Command4_Click()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
txt (False)
cmd (True)
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.Update
Adodc1.Refresh
txt (False)
cmd (True)
Adodc1.Refresh
End Sub

Private Sub Command6_Click()
Unload Me
End Sub
Private Function txt(A As Boolean)
Text1.Enabled = A
Text2.Enabled = A
Text3.Enabled = A
Text4.Enabled = A
Text5.Enabled = A
Text6.Enabled = A
Text7.Enabled = A
End Function

Private Function cmd(A As Boolean)
Command1.Enabled = A
Command2.Enabled = A
Command3.Enabled = A
Command4.Enabled = Not A
Command5.Enabled = Not A
End Function
