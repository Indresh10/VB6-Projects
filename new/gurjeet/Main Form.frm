VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIFrm 
   BackColor       =   &H8000000C&
   Caption         =   "Library"
   ClientHeight    =   8550
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13455
   Icon            =   "Main Form.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   8175
      Width           =   13455
      _ExtentX        =   23733
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Current User : "
            TextSave        =   "Current User : "
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            TextSave        =   "06-01-2020"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            TextSave        =   "05:21 PM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Picture         =   "Main Form.frx":0ECA
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Pct1 
      Align           =   1  'Align Top
      Height          =   7935
      Left            =   0
      ScaleHeight     =   7875
      ScaleWidth      =   13395
      TabIndex        =   15
      Top             =   0
      Width           =   13455
      Begin VB.CommandButton CmdUAcc 
         Caption         =   "User Account"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4920
         TabIndex        =   1
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "E&xit"
         BeginProperty Font 
            Name            =   "Palatino Linotype"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5040
         TabIndex        =   12
         Top             =   6000
         Width           =   2775
      End
      Begin VB.Frame FramMbr 
         Caption         =   "Member Operation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2655
         Left            =   630
         TabIndex        =   2
         Top             =   2880
         Width           =   3255
         Begin VB.CommandButton CmdMbrRpt 
            Caption         =   "Member Report"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   4
            Top             =   1710
            Width           =   2775
         End
         Begin VB.CommandButton CmdMbrEntry 
            Caption         =   "Member Operation"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   3
            Top             =   645
            Width           =   2775
         End
      End
      Begin VB.Frame FramBk 
         Caption         =   "Book Operation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2655
         Left            =   8280
         TabIndex        =   9
         Top             =   2880
         Width           =   3255
         Begin VB.CommandButton CmdBkRpt 
            Caption         =   "Book/CD Report"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   11
            Top             =   1710
            Width           =   2775
         End
         Begin VB.CommandButton CmdBkEntry 
            Caption         =   "Book Operation"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   10
            Top             =   645
            Width           =   2775
         End
      End
      Begin VB.Frame FramIsu 
         Caption         =   "Book Issue / submit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   3345
         Left            =   4440
         TabIndex        =   5
         Top             =   2550
         Width           =   3255
         Begin VB.CommandButton CmdIsuRpt 
            Caption         =   "Issue Report"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   8
            Top             =   2445
            Width           =   2775
         End
         Begin VB.CommandButton CmdIsuDtl 
            Caption         =   "Book Issue Detail"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   6
            Top             =   630
            Width           =   2775
         End
         Begin VB.CommandButton CmdBkSubISu 
            Caption         =   "Issue/Submit Book"
            BeginProperty Font 
               Name            =   "Palatino Linotype"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   7
            Top             =   1515
            Width           =   2775
         End
      End
      Begin VB.Label LblTask 
         Caption         =   "Pick a task you want"
         BeginProperty Font 
            Name            =   "Sylfaen"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   3120
         TabIndex        =   0
         Top             =   120
         Width           =   6255
      End
      Begin VB.Label LblClose 
         AutoSize        =   -1  'True
         Caption         =   "Close Selection Form"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9270
         MouseIcon       =   "Main Form.frx":1DA4
         MousePointer    =   99  'Custom
         TabIndex        =   13
         Top             =   7140
         Width           =   2235
      End
   End
   Begin VB.Menu MnuMstr 
      Caption         =   "&Master"
      Begin VB.Menu MnuMstrSelection 
         Caption         =   "Master &Selection"
      End
   End
   Begin VB.Menu MnuMbr 
      Caption         =   "Mem&ber"
      Begin VB.Menu MnuMbrOpr 
         Caption         =   "&Member Operations"
         Shortcut        =   ^M
      End
   End
   Begin VB.Menu MnuBk 
      Caption         =   "B&ook"
      Begin VB.Menu MnuBkOpr 
         Caption         =   "Book &Operations"
         Shortcut        =   ^B
      End
      Begin VB.Menu MnuBkSptr1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuBkIsuSub 
         Caption         =   "&Issue/Submit Book"
         Shortcut        =   ^I
      End
      Begin VB.Menu MnuBkIsuDtl 
         Caption         =   "Issue &Detail"
      End
   End
   Begin VB.Menu MnuUmg 
      Caption         =   "&User Manager"
      Begin VB.Menu MnuUmgAcc 
         Caption         =   "User &Account"
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu MnuRpt 
      Caption         =   "&Report"
      Begin VB.Menu mnuIsuRpt 
         Caption         =   "&Issue Report"
      End
      Begin VB.Menu MnuMbrRpt 
         Caption         =   "&Member Report"
      End
      Begin VB.Menu MnuBkRpt 
         Caption         =   "&Book Report"
      End
      Begin VB.Menu MnuCdRpt 
         Caption         =   "&CD Report"
      End
   End
   Begin VB.Menu MnuWin 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu MnuWinCscd 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu MnuWinVrtl 
         Caption         =   "&Verticle"
      End
      Begin VB.Menu MnuWinHrz 
         Caption         =   "&Horizontal"
      End
      Begin VB.Menu MnuWinSptr1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuWinClose 
         Caption         =   "Close &All Forms"
      End
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "&About"
      Begin VB.Menu MnuAbtLib 
         Caption         =   "About &Library Management"
      End
   End
End
Attribute VB_Name = "MDIFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim FL As String 'TO STORE FILE NAME
Dim rpt As String



Private Sub CmdBkEntry_Click()
    Call MnuBkOpr_Click
End Sub

Private Sub CmdBkRpt_Click()
    Dim str As String
    str = InputBox("Enter BOOK for Book report and CD for Cd report", _
        "Report Creation", "BOOK")
        
    If str = "BOOK" Then
        
        Call MnuBkRpt_Click
    ElseIf str = "CD" Then
        Call MnuCdRpt_Click
    Else
        MsgBox "Invalid input.", vbCritical, "Report Creation"
    End If
End Sub

Private Sub CmdBkSubISu_Click()
    Call MnuBkIsuSub_Click
End Sub

Private Sub CmdExit_Click()
    End
End Sub

Private Sub CmdIsuDtl_Click()
    Call MnuBkIsuDtl_Click
End Sub

Private Sub CmdIsuRpt_Click()
    Call mnuIsuRpt_Click
End Sub

Private Sub CmdMbrEntry_Click()
    Call MnuMbrOpr_Click
End Sub

Private Sub CmdMbrRpt_Click()
    Call MnuMbrRpt_Click
End Sub

Private Sub CmdUAcc_Click()
    Call MnuUmgAcc_Click
End Sub

Private Sub LblClose_Click()
    Pct1.Visible = False
End Sub

Private Sub MDIForm_Load()
    'CHECK USER TYPE
    If userNm = "LIBRARY" Then
         MnuUmg.Enabled = False
         CmdUAcc.Enabled = False
    End If
    
    If userType = "L" Then
        MnuBkIsuSub.Enabled = False
        MnuRpt.Enabled = False
        
        CmdBkSubISu.Enabled = False
    End If
End Sub

Private Sub MDIForm_Resize()
    'RESIZE STATUS BAR
    If Me.Width > 1000 And Me.Height > 1000 Then
        StatusBar1.Panels(1).Width = Me.ScaleWidth * 0.5
        StatusBar1.Panels(2).Width = Me.ScaleWidth * 0.11
        StatusBar1.Panels(3).Width = Me.ScaleWidth * 0.11
        StatusBar1.Panels(4).Width = Me.ScaleWidth * 0.11
        StatusBar1.Panels(5).Width = Me.ScaleWidth * 0.11
        StatusBar1.Panels(6).Width = Me.ScaleWidth * 0.05
        StatusBar1.Panels(1) = "Current User : " & userNm & "(" & userType & ")"
    End If
    
    'ARRANGE PICTURE BOX AND OTHER COMMAND BUTTONS
    Pct1.Height = Me.Height
    
    If Me.Height >= 8100 And Me.Width >= 11500 Then
        LblTask.Left = Me.ScaleWidth / 2 - LblTask.Width / 2    'MAKE LABLE TO CENTER

        'SET ALL COMMAND BUTTONS AND FRAME
        CmdUAcc.Left = Me.ScaleWidth / 2 - CmdUAcc.Width / 2 'SET COMMAND BUTTON TO CENTER
        
        
        FramIsu.Left = Me.ScaleWidth / 2 - FramIsu.Width / 2
        FramMbr.Left = FramIsu.Left - FramMbr.Width - 500
        FramBk.Left = FramIsu.Left + FramIsu.Width + 500
        
        CmdExit.Left = Me.ScaleWidth / 2 - CmdExit.Width / 2 'SET COMMAND BUTTON TO CENTER

        LblClose.Top = Me.Height - 1500
        LblClose.Left = Me.ScaleWidth - 2500
    End If
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    End
End Sub

Private Sub MnuAbtLib_Click()
    Pct1.Visible = False
    frmAbout.Show vbModal
End Sub


Private Sub MnuBkIsuDtl_Click()
    Pct1.Visible = False
    FrmIsuDtl.Show
End Sub

Private Sub MnuBkIsuSub_Click()
    Pct1.Visible = False
    FrmBookIsu.Show
End Sub

Private Sub MnuBkOpr_Click()
    Pct1.Visible = False
    frmBkEntry.Show
End Sub

Private Sub MnuBkRpt_Click()
    
    Call BookCdReport("BOOK")   'GENERATE REPORT

End Sub

Private Sub MnuCdRpt_Click()
    
    Call BookCdReport("CD")   'GENERATE REPORT

End Sub

Private Sub mnuIsuRpt_Click()
    Pct1.Visible = False
    
    Report = "I" 'I means Issue Report
    FrmRpt.Show vbModal
End Sub


Private Sub MnuMbrOpr_Click()
    Pct1.Visible = False
    FrmMember.Show
End Sub

Private Sub MnuMbrRpt_Click()
    Pct1.Visible = False
    Report = "M" 'M means Member Report
    FrmRpt.Show vbModal
End Sub

Private Sub MnuMstrSelection_Click()
    Pct1.Visible = True
End Sub

Private Sub MnuUmgAcc_Click()
    Pct1.Visible = False
    FrmUserMng.Show vbModal
End Sub

Private Sub MnuWinClose_Click()
    Do While Forms.Count - 1 > 0
        Unload Me.ActiveForm
    Loop
End Sub

Private Sub MnuWinCscd_Click()
    Arrange vbCascade
End Sub

Private Sub MnuWinHrz_Click()
    Arrange vbHorizontal
End Sub

Private Sub MnuWinVrtl_Click()
    Arrange vbVertical
End Sub


'==================================================================================
'GENERATE REPORT FOR BOOK/CD
Private Sub BookCdReport(typ As String)
    Set rs = New Recordset
    
    If typ = "BOOK" Then
        rs.Open "SELECT Code,Title,Author,Price,Qty FROM Book_Mast WHERE Code like 'B%'", conn, adOpenStatic, adLockReadOnly
    Else
        rs.Open "SELECT Code,Title,Author,Price,Qty FROM Book_Mast WHERE Code like 'C%'", conn, adOpenStatic, adLockReadOnly
    End If
        
    'WHEN NO RECORD EXIST
    If rs.RecordCount = 0 Then
        rs.Close
        MsgBox "No record is found.", vbInformation, "Member Report"
        Exit Sub
    End If
        
    'CREATE REPORT
    'OPEN FILE
    FL = typ & "_" & Format(Date, "dd-mm-yyyy")
    Open App.Path & "\Reports\" & FL & ".txt" For Output As #1
        
    Print #1, ""
    Print #1, "--------------------------------------------------------------------------------"
    
    If typ = "BOOK" Then
        Print #1, "---------------------------- B O O K S  R E P O R T ----------------------------"
    Else
        Print #1, "------------------------------- C D  R E P O R T -------------------------------"
    End If
    
    Print #1, "--------------------------------------------------------------------------------"
    Print #1, ""
    Print #1, " Date : " & Format(Date, "dd-mm-yyyy")
    Print #1, ""
    Print #1, "--------------------------------------------------------------------------------"
    Print #1, " CODE    TITLE                          AUTHOR                 PRICE   QUANTITY "
    Print #1, "--------------------------------------------------------------------------------"
        
    rs.MoveFirst
    Do While Not rs.EOF
        Print #1, " " & rs!Code & "  " & _
                rs!title & Space(31 - Len(rs!title)) & _
                rs!Author & Space(22 - Len(rs!Author)) & _
                Space(6 - Len(rs!Price)) & rs!Price & _
                Space(11 - Len(rs!qty)) & rs!qty
        Print #1, ""
        rs.MoveNext
    Loop
    rs.Close
        
    Close #1
    MsgBox FL & ".txt created successfully.", vbInformation, "Member Report"
        
    Shell App.Path & "\Reports\wordpad.exe " & App.Path & "\Reports\" & FL & ".txt", vbMaximizedFocus

End Sub
