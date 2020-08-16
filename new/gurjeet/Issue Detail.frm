VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmIsuDtl 
   Caption         =   "Issue Detail"
   ClientHeight    =   6435
   ClientLeft      =   1815
   ClientTop       =   900
   ClientWidth     =   7710
   Icon            =   "Issue Detail.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleWidth      =   7710
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdIsu 
      Caption         =   "Issue / Submit"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   495
      TabIndex        =   1
      Top             =   5160
      Width           =   1980
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5985
      TabIndex        =   2
      Top             =   5160
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid Msfg1 
      Height          =   3975
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   7
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
   Begin VB.Label LblIsuDtl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Issue Detail"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400040&
      Height          =   480
      Left            =   2760
      TabIndex        =   3
      Top             =   0
      Width           =   1950
   End
End
Attribute VB_Name = "FrmIsuDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_isu As New Recordset
Dim i As Integer


Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdIsu_Click()
    FrmBookIsu.Show
End Sub

Private Sub Form_GotFocus()
    Call Form_Load
End Sub

Private Sub Form_Load()
    If userType = "L" Then
        CmdIsu.Enabled = False
    End If
    
    LblIsuDtl.Left = Me.ScaleWidth / 2 - LblIsuDtl.Width / 2
    If Msfg1.Height > 1 Then
        Msfg1.Width = FrmIsuDtl.ScaleWidth - 500
    End If
    
    CmdCancel.Left = Me.ScaleWidth - CmdCancel.Width - 500
    CmdCancel.Top = Msfg1.Height + 1200
    
    'OPEN CONNECTION
    Set rs_isu = New Recordset
    rs_isu.Open "SELECT * FROM Issue_Mast", conn, adOpenStatic, adLockReadOnly
    
    Msfg1.Rows = rs_isu.RecordCount + 1
    
    Msfg1.FormatString = "No.  | Member Code | Class  | Year  | Book Code | Issued Date | Last Submit Date"
    
    Dim temp
    If rs_isu.RecordCount > 0 Then
        rs_isu.MoveFirst
        For i = 1 To rs_isu.RecordCount
            Msfg1.TextMatrix(i, 0) = i
            Msfg1.TextMatrix(i, 1) = rs_isu.Fields(0)
            Msfg1.TextMatrix(i, 2) = rs_isu.Fields(1)
            Msfg1.TextMatrix(i, 3) = rs_isu.Fields(2)
            Msfg1.TextMatrix(i, 4) = rs_isu.Fields(3)
            Msfg1.TextMatrix(i, 5) = Format(rs_isu.Fields(4), "dd-mmm-yyyy")
            Msfg1.TextMatrix(i, 6) = Format(rs_isu.Fields(5), "dd-mmm-yyyy")
            
            rs_isu.MoveNext
        Next
    End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 4365 And Me.Width > 8500 Then
        'PUT LABEL IN MIDDLE OF FORM
        LblIsuDtl.Left = Me.ScaleWidth / 2 - LblIsuDtl.Width / 2
        If FrmIsuDtl.ScaleHeight - 2000 > 1 Then
            Msfg1.Height = FrmIsuDtl.ScaleHeight - 2000
            Msfg1.Width = FrmIsuDtl.ScaleWidth - 500
        End If

        CmdIsu.Top = Msfg1.Height + 1200
        
        CmdCancel.Left = Me.ScaleWidth - CmdCancel.Width - 500
        CmdCancel.Top = Msfg1.Height + 1200
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Forms.Count = 2 Then
        MDIFrm.Pct1.Visible = True
    End If
End Sub
