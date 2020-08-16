VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMbrDtl 
   Caption         =   "Member Detail"
   ClientHeight    =   6210
   ClientLeft      =   2460
   ClientTop       =   945
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   6975
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   810
      TabIndex        =   6
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
         ItemData        =   "Member_detail.frx":0000
         Left            =   3840
         List            =   "Member_detail.frx":001F
         Style           =   2  'Dropdown List
         TabIndex        =   8
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
         ItemData        =   "Member_detail.frx":0053
         Left            =   990
         List            =   "Member_detail.frx":006F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   190
         Width           =   1215
      End
      Begin VB.Label LblYear 
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
         Left            =   3240
         TabIndex        =   10
         Top             =   250
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
         Left            =   270
         TabIndex        =   9
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   5280
      Width           =   1215
   End
   Begin VB.CommandButton CmdMbrEntry 
      Caption         =   "&Enter Member"
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
      TabIndex        =   2
      Top             =   5280
      Width           =   1575
   End
   Begin VB.CommandButton CmdMbrDlt 
      Caption         =   "&Delete Member"
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
      Left            =   2280
      TabIndex        =   3
      Top             =   5280
      Width           =   1455
   End
   Begin MSFlexGridLib.MSFlexGrid MsfgMbr 
      Height          =   3975
      Left            =   240
      TabIndex        =   1
      Top             =   1095
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   8
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
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
Option Explicit

Dim conn As New Connection
Dim rs_mem As New Recordset
Dim i As Integer

Private Sub CmbClass_Click()
    Call fillYear(Me)
    CmbClassYear.Text = CmbClassYear.List(0)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdMbrDlt_Click()
    FrmMbrDelt.Show vbModal
End Sub

Private Sub CmdMbrEntry_Click()
    FrmMmbrEntry.Show vbModal
End Sub

Private Sub Form_Load()
    
    
    LblMbrDtl.Left = Me.ScaleWidth / 2 - LblMbrDtl.Width / 2
    Frame1.Left = Me.ScaleWidth / 2 - Frame1.Width / 2
    
    Me.Height = 3000
    If MsfgMbr.Height > 1 Then
        MsfgMbr.Height = FrmMbrDtl.ScaleHeight - 2000
        MsfgMbr.Width = FrmMbrDtl.ScaleWidth - 500
    End If
    CmdMbrEntry.Top = MsfgMbr.Height + 1200
    CmdMbrDlt.Top = MsfgMbr.Height + 1200
    CmdRef.Top = MsfgMbr.Height + 1200
    CmdCancel.Left = Me.ScaleWidth - CmdCancel.Width - 500
    CmdCancel.Top = MsfgMbr.Height + 1200
    
    
    MsfgMbr.FormatString = "No. | Code   | Name                                             |" & _
            " Join Date  | City          | Contacet No. | Fee    | Fine  "
    
'    If rs_mem.RecordCount <> 0 Then
'        rs_mem.MoveFirst
'        For i = 1 To rs_mem.RecordCount
'            MsfgMbr.TextMatrix(i, 0) = rs_mem.Fields(0)
'            MsfgMbr.TextMatrix(i, 1) = rs_mem.Fields(1)
'            MsfgMbr.TextMatrix(i, 2) = rs_mem.Fields(2)
'            MsfgMbr.TextMatrix(i, 3) = rs_mem.Fields(3)
'            MsfgMbr.TextMatrix(i, 4) = rs_mem.Fields(4)
'            MsfgMbr.TextMatrix(i, 5) = rs_mem.Fields(5)
'            MsfgMbr.TextMatrix(i, 6) = rs_mem.Fields(6)
'            MsfgMbr.TextMatrix(i, 7) = rs_mem.Fields(7)
'            MsfgMbr.TextMatrix(i, 8) = rs_mem.Fields(8)
'            MsfgMbr.TextMatrix(i, 9) = rs_mem.Fields(9)
''            MsfgMbr.TextMatrix(i, 8) = rs_fee.Fields(1)
''            MsfgMbr.TextMatrix(i, 9) = rs_fee.Fields(2)
'            rs_mem.MoveNext
''            rs_fee.MoveNext
'        Next
'    Else
'        MsgBox "No member is entered", vbInformation, "Member Detail"
'        Exit Sub
'    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 4365 And Me.Width > 7320 Then
        LblMbrDtl.Left = Me.ScaleWidth / 2 - LblMbrDtl.Width / 2
        Frame1.Left = Me.ScaleWidth / 2 - Frame1.Width / 2
        
        If FrmMbrDtl.ScaleHeight - 2000 > 1 Then
            MsfgMbr.Height = FrmMbrDtl.ScaleHeight - 2000
            MsfgMbr.Width = FrmMbrDtl.ScaleWidth - 500
        End If
        CmdMbrEntry.Top = MsfgMbr.Height + 1200
        CmdMbrDlt.Top = MsfgMbr.Height + 1200
        CmdRef.Top = MsfgMbr.Height + 1200
        CmdCancel.Left = Me.ScaleWidth - CmdCancel.Width - 500
        CmdCancel.Top = MsfgMbr.Height + 1200
    End If
End Sub

