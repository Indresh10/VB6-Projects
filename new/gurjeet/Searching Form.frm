VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmSearch 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Book/Cd Searching"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   7590
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSelect 
      Caption         =   "&Select searched record"
      Default         =   -1  'True
      Height          =   375
      Left            =   480
      TabIndex        =   12
      Top             =   6120
      Width           =   2175
   End
   Begin VB.TextBox TxtSearch 
      Height          =   375
      Left            =   1800
      TabIndex        =   9
      Top             =   2550
      Width           =   3855
   End
   Begin VB.CommandButton CmdSearch 
      Caption         =   "&Search"
      Height          =   375
      Left            =   5880
      TabIndex        =   10
      Top             =   2550
      Width           =   1095
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   375
      Left            =   5880
      TabIndex        =   13
      Top             =   6120
      Width           =   1095
   End
   Begin MSFlexGridLib.MSFlexGrid MsfgSearch 
      Height          =   2535
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.Frame FremType 
      Caption         =   "Search &Type"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2895
      Begin VB.OptionButton OptCd 
         Caption         =   "CD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   960
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.OptionButton OptBook 
         Caption         =   "Book"
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
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   855
      End
   End
   Begin VB.Frame FremCategory 
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
      Height          =   1455
      Left            =   3240
      TabIndex        =   3
      Top             =   720
      Width           =   4215
      Begin VB.OptionButton OptCode 
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
         Left            =   360
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptTitle 
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
         Left            =   360
         TabIndex        =   6
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton OptAuther 
         Caption         =   "Auther"
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
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton OptPublisher 
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
         Left            =   2760
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
   End
   Begin VB.Label LblSearch 
      AutoSize        =   -1  'True
      Caption         =   "Search word :"
      Height          =   195
      Left            =   600
      TabIndex        =   8
      Top             =   2640
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Book/CD Searching"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   2280
      TabIndex        =   14
      Top             =   120
      Width           =   3450
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   7695
   End
End
Attribute VB_Name = "FrmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_srch As New Recordset

Private Sub CmdExit_Click()
    Unload Me
    frmBkEntry.Show
End Sub

Private Sub CmdSearch_Click()
    If Not (OptBook.Value Or OptCd.Value) Then
        MsgBox "select type"
    End If
    
    If Not (OptCode.Value Or OptAuther.Value Or OptTitle.Value Or OptPublisher.Value) Then
        MsgBox "select category"
    End If
End Sub

Private Sub Form_Load()
    Dim lbl As String
    OptBook.Value = True
    'SET SEARCH LABEL
    Call searchLabel(Me)
    
End Sub

Private Sub Form_Resize()
    Shape1.Width = Me.ScaleWidth
    Label1.Left = Me.ScaleWidth / 2 - Label1.Width / 2
    
End Sub

Private Sub OptAuther_Click()
    Call searchLabel(Me) 'SET SEARCH LABEL
    Call fillGrid(Me, "Auther", rs_srch) 'FILL FLEX GRID
End Sub

Private Sub OptBook_Click()
    Call searchLabel(Me)    'SET SEARCH LABEL CAPTION
    OptCode.Value = True
    Call OptCode_Click
End Sub

Private Sub OptCd_Click()
    Call searchLabel(Me)    'SET SEARCH LABEL CAPTION
    OptCode.Value = True
    Call OptCode_Click
End Sub

Private Sub OptCode_Click()
    Call searchLabel(Me)
    Call fillGrid(Me, "Code", rs_srch)
    
End Sub

Private Sub OptPublisher_Click()
    Call searchLabel(Me)
End Sub

Private Sub OptTitle_Click()
    Call searchLabel(Me)
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    KeyAscii = upper(KeyAscii)
End Sub

'=======================================================
'TO SEARCH BOOK/CD RECORD
Public Sub searchRecord(ByRef rs As Recordset)
    rs.Find "code = '" & MsfgSearch.TextMatrix(MsfgSearch.Row, 1) & "'"
    rs.Bookmark
End Sub

