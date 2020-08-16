VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMember 
   Caption         =   "Member Operations"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   Icon            =   "Member_Form.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   8490
   ScaleWidth      =   11880
   WindowState     =   2  'Maximized
   Begin VB.Frame FremCategory 
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
      TabIndex        =   42
      Top             =   720
      Width           =   5445
      Begin VB.TextBox TxtSearch 
         Height          =   375
         Left            =   2670
         MaxLength       =   15
         TabIndex        =   46
         Top             =   720
         Width           =   2415
      End
      Begin VB.ComboBox CmbSearch 
         Height          =   315
         ItemData        =   "Member_Form.frx":08CA
         Left            =   2670
         List            =   "Member_Form.frx":08DD
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   285
         Width           =   2415
      End
      Begin MSFlexGridLib.MSFlexGrid MsfgSearch 
         Height          =   5340
         Left            =   75
         TabIndex        =   47
         Top             =   1245
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   9419
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         Left            =   360
         TabIndex        =   45
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label LblSearch 
         AutoSize        =   -1  'True
         Caption         =   "Select Search Caregory :"
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
         Left            =   375
         TabIndex        =   43
         Top             =   300
         Width           =   2235
      End
   End
   Begin VB.TextBox TxtLast 
      Height          =   375
      Left            =   4290
      MaxLength       =   15
      TabIndex        =   14
      Top             =   1920
      Width           =   1845
   End
   Begin VB.TextBox TxtFirst 
      Height          =   375
      Left            =   2520
      MaxLength       =   15
      TabIndex        =   13
      Top             =   1920
      Width           =   1770
   End
   Begin VB.TextBox TxtCity 
      Height          =   375
      Left            =   1200
      MaxLength       =   20
      TabIndex        =   23
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox TxtAddress 
      Height          =   1215
      Left            =   1200
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   21
      Top             =   2880
      Width           =   4935
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "&First"
      Height          =   615
      Left            =   1575
      Picture         =   "Member_Form.frx":0906
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "First Record"
      Top             =   6030
      Width           =   735
   End
   Begin VB.CommandButton CmdPrv 
      Caption         =   "&Previous"
      Height          =   615
      Left            =   2415
      Picture         =   "Member_Form.frx":0A08
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Previous Record"
      Top             =   6030
      Width           =   735
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next"
      Height          =   615
      Left            =   3255
      Picture         =   "Member_Form.frx":0B0A
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Next Record"
      Top             =   6030
      Width           =   735
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "&Last"
      Height          =   615
      Left            =   4095
      Picture         =   "Member_Form.frx":0C0C
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Last Record"
      Top             =   6030
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   630
      Left            =   735
      Picture         =   "Member_Form.frx":0D0E
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Add Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   630
      Left            =   1575
      Picture         =   "Member_Form.frx":1050
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Edit Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdTransfer 
      Caption         =   "Transfer"
      Height          =   630
      Left            =   4095
      Picture         =   "Member_Form.frx":1392
      Style           =   1  'Graphical
      TabIndex        =   40
      ToolTipText     =   "Search Record"
      Top             =   6735
      Width           =   735
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "&Delete"
      Height          =   630
      Left            =   2415
      Picture         =   "Member_Form.frx":16D4
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Delete Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   630
      Left            =   3255
      Picture         =   "Member_Form.frx":1A16
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Save Record"
      Top             =   6750
      Width           =   735
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   630
      Left            =   4935
      Picture         =   "Member_Form.frx":2080
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Exit"
      Top             =   6750
      Width           =   735
   End
   Begin VB.Frame FremPerInfo 
      Caption         =   "Personal &Information"
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
      TabIndex        =   26
      Top             =   4680
      Width           =   5895
      Begin VB.OptionButton OptFemale 
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
         TabIndex        =   31
         Top             =   540
         Width           =   1095
      End
      Begin VB.OptionButton OptMale 
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
         TabIndex        =   30
         Top             =   180
         Width           =   975
      End
      Begin VB.TextBox TxtContact 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   1440
         MaxLength       =   13
         TabIndex        =   28
         Top             =   420
         Width           =   1815
      End
      Begin VB.Label LblGender 
         AutoSize        =   -1  'True
         Caption         =   "&Gender :"
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
         TabIndex        =   29
         Top             =   300
         Width           =   765
      End
      Begin VB.Label LblContact 
         AutoSize        =   -1  'True
         Caption         =   "C&ontect No :"
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
         TabIndex        =   27
         Top             =   480
         Width           =   1080
      End
   End
   Begin VB.TextBox TxtFee 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4920
      MaxLength       =   5
      TabIndex        =   25
      Top             =   4200
      Width           =   1215
   End
   Begin VB.ComboBox CmbDay 
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
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   2400
      Width           =   735
   End
   Begin VB.ComboBox CmbMonth 
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
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   2400
      Width           =   855
   End
   Begin VB.ComboBox CmbYear 
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
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   2400
      Width           =   975
   End
   Begin VB.TextBox TxtSurname 
      Height          =   375
      Left            =   1200
      MaxLength       =   15
      TabIndex        =   12
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox TxtCode 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1200
      MaxLength       =   6
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1275
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   660
      Left            =   480
      TabIndex        =   1
      Top             =   495
      Width           =   5415
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
         ItemData        =   "Member_Form.frx":23BE
         Left            =   990
         List            =   "Member_Form.frx":23DA
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   190
         Width           =   1215
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
         ItemData        =   "Member_Form.frx":240D
         Left            =   3840
         List            =   "Member_Form.frx":242C
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   195
         Width           =   1215
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
         TabIndex        =   2
         Top             =   255
         Width           =   600
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
         TabIndex        =   4
         Top             =   250
         Width           =   525
      End
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
      Left            =   3300
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
      Width           =   11880
   End
   Begin VB.Label LblLast 
      AutoSize        =   -1  'True
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
      Left            =   4560
      TabIndex        =   10
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label LblFirst 
      AutoSize        =   -1  'True
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
      Left            =   2640
      TabIndex        =   9
      Top             =   1680
      Width           =   1350
   End
   Begin VB.Label LblSurname 
      AutoSize        =   -1  'True
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
      Left            =   1440
      TabIndex        =   8
      Top             =   1680
      Width           =   810
   End
   Begin VB.Label LblFee 
      AutoSize        =   -1  'True
      Caption         =   "Member &Fee :"
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
      TabIndex        =   24
      Top             =   4260
      Width           =   1245
   End
   Begin VB.Label LblCity 
      AutoSize        =   -1  'True
      Caption         =   "Ci&ty :"
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
      TabIndex        =   22
      Top             =   4260
      Width           =   420
   End
   Begin VB.Label LblAddress 
      AutoSize        =   -1  'True
      Caption         =   "Addre&ss :"
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
      Top             =   2910
      Width           =   855
   End
   Begin VB.Label LblDtFrmt 
      AutoSize        =   -1  'True
      Caption         =   "DD-MM-YYYY"
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
      Left            =   3840
      TabIndex        =   19
      Top             =   2460
      Width           =   1290
   End
   Begin VB.Label LblJoinDate 
      AutoSize        =   -1  'True
      Caption         =   "&Join Date :"
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
      Top             =   2460
      Width           =   945
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Na&me :"
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
      Top             =   1980
      Width           =   645
   End
   Begin VB.Label LblCode 
      AutoSize        =   -1  'True
      Caption         =   "Co&de :"
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
      Top             =   1335
      Width           =   585
   End
End
Attribute VB_Name = "FrmMember"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs_mbr As New ADODB.Recordset
Dim rs_temp As New ADODB.Recordset
Dim rs_isu As New ADODB.Recordset

Dim cmd As String

Private Sub CmbClass_Click()
    Dim i As Integer
    
    Class = CmbClass.Text
    
    Call fillYear(Me) 'SELECT YEAR
    
    CmbClassYear.Text = CmbClassYear.List(0)
End Sub


Private Sub CmbClassYear_Click()
        
    Yer = CmbClassYear.Text
    
    Call Member.controlEnable(Me, False)

    If rs_mbr.State = 1 Then rs_mbr.Close
    rs_mbr.Open "select * from Mbr_Mast where [crs]='" & CmbClass.Text & _
            "' and [Yer]='" & CmbClassYear.Text & "' ORDER BY Code", conn, adOpenStatic, adLockPessimistic
            
    Call Member.clearControl(Me) 'SET DEFAULT CONTROLS
    If rs_mbr.RecordCount <> 0 Then
        Call Book.enableCommand(Me) 'ENABLE COMMAND BTNS
        Call memberData(Me, rs_mbr) 'RETRIVE DATA
    Else
        Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
        CmdAdd.Enabled = True
    End If
    
    'CHECK USER TYPE
    If userType = "L" Then
        CmdAdd.Enabled = False
        CmdEdit.Enabled = False
        CmdDel.Enabled = False
        CmdSave.Enabled = False
        CmdTransfer.Enabled = False
    End If
End Sub

Private Sub CmbMonth_Click()
    Dim i As Integer
    CmbDay.Clear
    For i = 1 To daysOfMonth(Val(CmbMonth.Text), Val(CmbYear.Text))
        CmbDay.AddItem i
    Next i
    CmbDay.Text = Day(Date)
End Sub

Private Sub CmbSearch_Click()
    Call fillMbrGrid(Me, CmbClass.Text, CmbClassYear.Text, CmbSearch.Text)
End Sub

Private Sub CmbYear_Click()
        Dim i As Integer
        CmbDay.Clear
        For i = 1 To daysOfMonth(Val(CmbMonth.Text), Val(CmbYear.Text))
            CmbDay.AddItem i
        Next i
        CmbDay.Text = Day(Date)
End Sub

Private Sub CmdAdd_Click()
    Dim rs_tmp As New ADODB.Recordset
    Set rs_tmp = New Recordset

    cmd = "Add"
    CmdExit.Caption = "&Cancel"
        
    'ENABLE ALL CONTROLS
    Call Member.controlEnable(Me, True)
    TxtCode.Locked = True
    CmdSave.Enabled = True      'ENABLE SAVE BUTTON
    CmbClass.Enabled = False    'DISABLE CLASS COMBO
    CmbClassYear.Enabled = False 'DISABLE YEAR COMBO
    FremCategory.Enabled = False 'DISABLE SEARCH FREAM
    Call Book.disableCommand(Me) 'DISABLE COMMAND BTNS
    CmdSave.Enabled = True
    
    'SET DEFALUT CONTROLS
    Call Member.clearControl(Me)
    
    'GENERATE NEXT CODE
    TxtCode.Text = Book.Next_Code(rs_mbr, "M") 'GENERATE NEXT CODE
    
    TxtSurname.SetFocus
End Sub

Private Sub CmdDel_Click()
    
    Set rs_temp = New Recordset
    rs_temp.Open "SELECT * FROM Issue_Mast WHERE [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "' AND [Mbr_No]='" & TxtCode & "' AND [Sub_Dt]='-'", conn, adOpenStatic, adLockReadOnly
    
    If rs_temp.RecordCount > 0 Then
        MsgBox "You can't delete this member. First Issue Book/CD.", vbInformation, "Member Deletion"
        Exit Sub
    End If


    If MsgBox("You want to delete this record?", vbInformation + vbYesNo, "Member deletion") = vbYes Then
        rs_mbr.Delete    'DELETE RECORD
        rs_mbr.Update    'UPDATE RECORD
        rs_mbr.MoveNext  'MOVE RECORDSET TO NEXT RECORD
        Call Member.fillMbrGrid(Me, Class, Yer, CmbSearch.Text)
        If rs_mbr.RecordCount = 0 Then
            Call Member.clearControl(Me)  'CLEAR TEXT BOXES
            Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
            CmdAdd.Enabled = True   'ENABLE ADD COMMAND BUTTONS
            Exit Sub
        Else
            If rs_mbr.EOF Then
                rs_mbr.MoveFirst
                Call Member.memberData(Me, rs_mbr) 'RETRIVE RECORD
                Exit Sub
            Else
                'Call CmbType_Click
                Call Member.memberData(Me, rs_mbr) 'RETRIVE RECORD
            End If
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    cmd = "Edit"
    CmdExit.Caption = "&Cancel"
        
    'ENABLE ALL CONTROLS
    Call Member.controlEnable(Me, True)
    FremCategory.Enabled = False 'DISABLE SEARCH FREAM
    Call Book.disableCommand(Me)
    CmdSave.Enabled = True  'ENABLE SAVE BTN
    
End Sub

Private Sub CmdExit_Click()
    If CmdExit.Caption = "&Cancel" Then
        CmdExit.Caption = "E&xit"
        
        'DISABLE ALL CONTROLS
        Call Member.controlEnable(Me, False)
        FremCategory.Enabled = True 'ENABLE SEARCH FREM
        
        Call Member.clearControl(Me)  'CLEAR CONTROLS
        If rs_mbr.RecordCount <> 0 Then
            rs_mbr.MoveFirst
            Call Book.enableCommand(Me) 'ENABLE COMMAND BUTTON
            Call Member.memberData(Me, rs_mbr) 'RETRIVE RECORD
        Else
            Call Book.disableCommand(Me) 'DISABLE COMMAND BTNS
            Call Member.clearControl(Me) 'CLEAR CONTROLS
            CmdAdd.Enabled = True
        End If
        Call Member.controlEnable(Me, False)  'LOCK TEXT BOXES
        CmdSave.Enabled = False
        CmbClass.Enabled = True 'ENABLE COURCE COMBO
        CmbClassYear.Enabled = True 'ENABLE YEAR COMBO

    ElseIf CmdExit.Caption = "E&xit" Then
        
        Unload Me
    End If
    
End Sub

Private Sub CmdFirst_Click()
    rs_mbr.MoveFirst 'MOVE RECORD TO FIRST
    Call Member.memberData(Me, rs_mbr) 'RETRIVE MEMBER DATA
End Sub

Private Sub CmdLast_Click()
    rs_mbr.MoveLast 'MOVE RECORD TO FIRST
    Call Member.memberData(Me, rs_mbr) 'RETRIVE MEMBER DATA
End Sub

Private Sub CmdNext_Click()
    rs_mbr.MovePrevious
    If rs_mbr.BOF Then
        rs_mbr.MoveLast
    End If
    Call Member.memberData(Me, rs_mbr)   'RETRIVE DATA
End Sub

Private Sub CmdPrv_Click()
    rs_mbr.MovePrevious
    If rs_mbr.BOF Then
        rs_mbr.MoveLast
    End If
    Call Member.memberData(Me, rs_mbr)   'RETRIVE DATA
End Sub

Private Sub CmdSave_Click()
    
    Dim dt As String, sex As String, Qry As String
    
    'VALIDATIONS
    If TxtCode = "" Or TxtSurname = "" Or TxtFirst = "" Or TxtLast = "" Or _
        TxtAddress = "" Or TxtCity = "" Or TxtFee = "" Then
            MsgBox "Enter all compulsory information.", vbInformation, "Member Entry"
            Exit Sub
    End If
        
        
    dt = CmbDay.Text & "/" & CmbMonth.Text & "/" & CmbYear.Text
    If OptMale.Value = True Then
        sex = "M"
    Else
        sex = "F"
    End If
    
    'ADD RECORD
    If cmd = "Add" Then
        
        Qry = "insert into Mbr_Mast values ('" & TxtCode & "','" & TxtSurname & "','" & _
            TxtFirst & "','" & TxtLast & "','" & dt & "','" & TxtAddress & "','" & _
            TxtCity & "','" & CmbClass.Text & "','" & CmbClassYear.Text & "','" & _
            TxtContact & "','" & sex & "'," & TxtFee & ",0)"
        
        conn.Execute Qry
        
        Call CmbClassYear_Click   'TO RETRIVE UPDATED DATA
        Call CmdExit_Click  'TO RESET CONTROLS

        MsgBox "Record added successfully.", vbInformation, "Member Entry"
        
    ElseIf cmd = "Edit" Then 'EDIT RECORD
        
        Qry = "update Mbr_Mast set [surname]='" & TxtSurname & "', [member]='" & _
            TxtFirst & "', [father]='" & TxtLast & "', [Join_Dt]='" & dt & "', [Address]='" & _
            TxtAddress & "',[City]='" & TxtCity & "', [Cnt_No]='" & _
            TxtContact & "',[Gender]='" & sex & "',[Fee]=" & TxtFee & " where [Code]='" & _
            TxtCode & "'" & " and [Crs]='" & Class & "' and [Yer]='" & Yer & "'"
        
            
        MsgBox Qry
        
        conn.Execute Qry
        
        Call CmdExit_Click  'TO RESET CONTROLS

    End If

End Sub

Private Sub CmdTransfer_Click()
    Unload Me
    FrmTransfer.Show vbModal
End Sub


Private Sub Form_Load()
    Dim i As Integer
    
    'DAY COMBO
    For i = 1 To 31
        CmbDay.AddItem i
    Next
    'MONTH COMBO
    For i = 1 To 12
        CmbMonth.AddItem i
    Next
    'YEAR COMBO
    For i = 1950 To 2050
        CmbYear.AddItem i
    Next
    
    CmbClass.Text = Class
    
    Me.MsfgSearch.FormatString = "No. |Code     |Name                                |Join Date   |City                   " & _
                                "|Contect No.    |Gender| Fine"
End Sub

Private Sub Form_Resize()
    If Me.Width > 6630 Then
        ShapLabel.Width = Me.ScaleWidth
        LblLabel.Left = ShapLabel.Width / 2 - LblLabel.Width / 2
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs_mbr.Close
    
    If Forms.Count = 2 Then
        MDIFrm.Pct1.Visible = True
    End If
End Sub

Private Sub MsfgSearch_Click()
    rs_mbr.MoveFirst
    rs_mbr.Find "Code = '" & MsfgSearch.TextMatrix(MsfgSearch.Row, 1) & "'"
    
    Call memberData(Me, rs_mbr) 'fill controls
End Sub


Private Sub MsfgSearch_RowColChange()
    rs_mbr.MoveFirst
    rs_mbr.Find "Code = '" & MsfgSearch.TextMatrix(MsfgSearch.Row, 1) & "'"
    
    Call memberData(Me, rs_mbr) 'fill controls
End Sub

Private Sub TxtAddress_GotFocus()
    Call Book.selectTxt(TxtAddress)
End Sub

Private Sub TxtAddress_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
        Exit Sub
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtCity_GotFocus()
    Call Book.selectTxt(TxtCity)
End Sub

Private Sub TxtCity_KeyPress(KeyAscii As Integer)
    KeyAscii = alpha(KeyAscii)
End Sub

Private Sub TxtCode_GotFocus()
    Call Book.selectTxt(TxtCode)
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtContact_GotFocus()
    Call Book.selectTxt(TxtContact)
End Sub

Private Sub TxtContact_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = 8
    ElseIf (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> Asc("-") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFee_GotFocus()
    Call Book.selectTxt(TxtFee)
End Sub

Private Sub TxtFee_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = 8
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtFirst_GotFocus()
    Call Book.selectTxt(TxtFirst)
End Sub

Private Sub TxtFirst_KeyPress(KeyAscii As Integer)
    KeyAscii = alpha(KeyAscii)
End Sub

Private Sub TxtLast_GotFocus()
    Call Book.selectTxt(TxtLast)
End Sub

Private Sub TxtLast_KeyPress(KeyAscii As Integer)
    KeyAscii = alpha(KeyAscii)
End Sub

Private Sub TxtSearch_Change()
    Set rs_temp = New Recordset
    
    rs_temp.Open "select * from Mbr_Mast where [Crs]='" & CmbClass.Text & "' and [Yer]='" & CmbClassYear.Text & "' and " & CmbSearch.Text & " like('" & TxtSearch & "%') order by " & CmbSearch, conn, adOpenStatic, adLockReadOnly
        
    If rs_temp.RecordCount = 0 Then
        MsfgSearch.Enabled = False
    Else
        MsfgSearch.Enabled = True
    End If
    
    Call fillMbrGrid1(Me, rs_temp) 'fill grid
End Sub


Private Sub TxtSearch_GotFocus()
    TxtSearch.Locked = False
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = upper(KeyAscii)
End Sub

Private Sub TxtSurname_GotFocus()
    Call Book.selectTxt(TxtSurname)
End Sub

Private Sub TxtSurname_KeyPress(KeyAscii As Integer)
    KeyAscii = Member.alpha(KeyAscii)
End Sub

'===========================================================
Private Sub fillMbrGrid1(Frm As Form, rs As Recordset)
    Dim r As Integer
    Frm.MsfgSearch.Cols = 8
    Frm.MsfgSearch.Rows = rs.RecordCount + 1
    
    If rs.RecordCount > 0 Then
    rs.MoveFirst
    For r = 1 To rs.RecordCount
        Frm.MsfgSearch.TextMatrix(r, 0) = r
        Frm.MsfgSearch.TextMatrix(r, 0) = r
                Frm.MsfgSearch.TextMatrix(r, 1) = rs.Fields(0)
                Frm.MsfgSearch.TextMatrix(r, 2) = rs.Fields(1) & " " & rs.Fields(2) & " " & rs.Fields(3)
                Frm.MsfgSearch.TextMatrix(r, 3) = Format(rs.Fields(4), "dd-mm-yyyy")
                Frm.MsfgSearch.TextMatrix(r, 4) = rs.Fields(6)
                Frm.MsfgSearch.TextMatrix(r, 5) = rs.Fields(9)
                Frm.MsfgSearch.TextMatrix(r, 6) = rs.Fields(10)
                Frm.MsfgSearch.TextMatrix(r, 7) = rs.Fields(12)
        rs.MoveNext
    Next
    End If
End Sub
