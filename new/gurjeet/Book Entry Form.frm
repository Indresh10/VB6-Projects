VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBkEntry 
   Caption         =   "Book Entry"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10650
   Icon            =   "Book Entry Form.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   10650
   WindowState     =   2  'Maximized
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
      Height          =   5895
      Left            =   6720
      TabIndex        =   32
      Top             =   1080
      Width           =   5055
      Begin VB.TextBox TxtSearch 
         Height          =   375
         Left            =   2070
         TabIndex        =   38
         Top             =   705
         Width           =   2400
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
         Left            =   3600
         TabIndex        =   36
         Top             =   390
         Width           =   1215
      End
      Begin VB.OptionButton OptAuthor 
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
         Left            =   1440
         TabIndex        =   34
         Top             =   390
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
         Left            =   2640
         TabIndex        =   35
         Top             =   390
         Width           =   975
      End
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
         TabIndex        =   33
         Top             =   390
         Width           =   975
      End
      Begin MSFlexGridLib.MSFlexGrid MsfgSearch 
         Height          =   4695
         Left            =   75
         TabIndex        =   39
         Top             =   1140
         Width           =   4920
         _ExtentX        =   8678
         _ExtentY        =   8281
         _Version        =   393216
         AllowUserResizing=   1
      End
      Begin VB.Label LblSearch 
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
         Left            =   555
         TabIndex        =   37
         Top             =   765
         Width           =   1485
      End
   End
   Begin VB.TextBox TxtAvlQty 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   375
      Left            =   5160
      MaxLength       =   6
      TabIndex        =   31
      Top             =   4680
      Width           =   1215
   End
   Begin VB.ComboBox CmbYear 
      Height          =   315
      Left            =   3120
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3720
      Width           =   855
   End
   Begin VB.ComboBox CmbMonth 
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox CmbDay 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   3720
      Width           =   735
   End
   Begin VB.TextBox TxtQty 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1680
      MaxLength       =   5
      TabIndex        =   18
      Top             =   4680
      Width           =   1215
   End
   Begin VB.TextBox TxtPrice 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   16
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox TxtAuther 
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox TxtTitle 
      Height          =   375
      Left            =   1680
      MaxLength       =   30
      TabIndex        =   5
      Top             =   2280
      Width           =   4695
   End
   Begin VB.TextBox TxtCode 
      Height          =   375
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox TxtPub 
      Height          =   375
      Left            =   1680
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3240
      Width           =   4695
   End
   Begin VB.TextBox TxtFrom 
      Height          =   375
      Left            =   1680
      MaxLength       =   50
      TabIndex        =   20
      Top             =   5160
      Width           =   4695
   End
   Begin VB.ComboBox CmbType 
      Height          =   315
      ItemData        =   "Book Entry Form.frx":030A
      Left            =   3240
      List            =   "Book Entry Form.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1095
      Width           =   1215
   End
   Begin VB.CommandButton CmdExit 
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   630
      Left            =   4980
      Picture         =   "Book Entry Form.frx":0322
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Exit"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   630
      Left            =   4140
      Picture         =   "Book Entry Form.frx":0660
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Save Record"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "&Delete"
      Height          =   630
      Left            =   3300
      Picture         =   "Book Entry Form.frx":0CCA
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Delete Record"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton CmdEdit 
      Caption         =   "&Edit"
      Height          =   630
      Left            =   2460
      Picture         =   "Book Entry Form.frx":100C
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Edit Record"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "&Add"
      Height          =   630
      Left            =   1620
      Picture         =   "Book Entry Form.frx":134E
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Add Record"
      Top             =   6480
      Width           =   735
   End
   Begin VB.CommandButton CmdLast 
      Caption         =   "&Last"
      Height          =   615
      Left            =   4560
      Picture         =   "Book Entry Form.frx":1690
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Last Record"
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton CmdNext 
      Caption         =   "&Next"
      Height          =   615
      Left            =   3720
      Picture         =   "Book Entry Form.frx":1792
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Next Record"
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton CmdPrv 
      Caption         =   "&Previous"
      Height          =   615
      Left            =   2880
      Picture         =   "Book Entry Form.frx":1894
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Previous Record"
      Top             =   5760
      Width           =   735
   End
   Begin VB.CommandButton CmdFirst 
      Caption         =   "&First"
      Height          =   615
      Left            =   2040
      Picture         =   "Book Entry Form.frx":1996
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "First Record"
      Top             =   5760
      Width           =   735
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BOOK/CD OPERATIONS"
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
      Left            =   3645
      TabIndex        =   40
      Top             =   45
      Width           =   4185
   End
   Begin VB.Shape ShapLabel 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   11895
   End
   Begin VB.Label LblAvlQty 
      AutoSize        =   -1  'True
      Caption         =   "Available Quantity :"
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
      TabIndex        =   30
      Top             =   4740
      Width           =   1710
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
      Left            =   4080
      TabIndex        =   14
      Top             =   3750
      Width           =   1290
   End
   Begin VB.Label LblQty 
      AutoSize        =   -1  'True
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
      TabIndex        =   17
      Top             =   4740
      Width           =   1320
   End
   Begin VB.Label LblPrice 
      AutoSize        =   -1  'True
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
      TabIndex        =   15
      Top             =   4200
      Width           =   555
   End
   Begin VB.Label LblAuther 
      AutoSize        =   -1  'True
      Caption         =   "Aut&her :"
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
      Top             =   2820
      Width           =   660
   End
   Begin VB.Label LblTitle 
      AutoSize        =   -1  'True
      Caption         =   "T&itle :"
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
      Top             =   2340
      Width           =   480
   End
   Begin VB.Label LblCode 
      AutoSize        =   -1  'True
      Caption         =   " Code :"
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
      Top             =   1860
      Width           =   630
   End
   Begin VB.Label LblPurDt 
      AutoSize        =   -1  'True
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
      TabIndex        =   10
      Top             =   3780
      Width           =   1425
   End
   Begin VB.Label LblPub 
      AutoSize        =   -1  'True
      Caption         =   "P&ublisher :"
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
      Top             =   3300
      Width           =   930
   End
   Begin VB.Label LblFrom 
      AutoSize        =   -1  'True
      Caption         =   "Fro&m (optional) :"
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
      TabIndex        =   19
      Top             =   5220
      Width           =   1440
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
      Left            =   2640
      TabIndex        =   0
      Top             =   1095
      Width           =   570
   End
End
Attribute VB_Name = "frmBkEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_bk As New ADODB.Recordset
Dim rs_temp As New ADODB.Recordset
Dim cmd As String

Private Sub CmbType_Click()

    bkType = CmbType.Text   'SET TYPE OF BOOK, CD OR BOOK/CD
    
    'CLOSE RECORDSET WHEN OPEN
    If rs_bk.State = 1 Then
        rs_bk.Close
    End If
    
    If CmbType.Text = "BOOK" Then
        rs_bk.Open "select * from Book_Mast where code like('B%')", conn, adOpenStatic, adLockPessimistic
        If rs_bk.RecordCount <> 0 Then
            Call Book.bookData(Me, rs_bk) 'RETRIVE DATA
            Call OptCode_Click  'CALL CLICK METHOD OF CODE OPTION BUTTON
            Call Book.enableCommand(Me)   'ENABLE COMMAND BUTTON
            CmdSave.Enabled = False
        Else
            Call Book.clearText(Me)  'CLEAR TEXTBOXES
            Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
            Call OptCode_Click
            CmdAdd.Enabled = True
        End If
    ElseIf CmbType.Text = "CD" Then
        rs_bk.Open "select * from Book_Mast where code like('C%')", conn, adOpenStatic, adLockPessimistic
        If rs_bk.RecordCount <> 0 Then
            Call Book.bookData(Me, rs_bk) 'RETRIVE DATA
            Call OptCode_Click  'CALL CLICK METHOD CODE OPTION BUTTON
            Call Book.enableCommand(Me)  'ENABLE COMMAND BUTTON
            CmdSave.Enabled = False
        Else
            Call Book.clearText(Me)  'CLEAR TEXTBOXES
            Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
            Call OptCode_Click  'CALL CLICK METHOD CODE OPTION BUTTON
            CmdAdd.Enabled = True
        End If
    End If
    
    'CHECK USER TYPE
    If userType = "L" Then
        CmdAdd.Enabled = False
        CmdEdit.Enabled = False
        CmdDel.Enabled = False
        CmdSave.Enabled = False
    End If
End Sub

Private Sub CmdAdd_Click()
    cmd = "Add"
    Call Book.enableText(Me) 'ENABLE TEXT BOXES
    CmbType.Enabled = False 'DISADBLE TYPE SELECTION
    Call Book.clearText(Me)  'CLEAR TEXT BOXES
    
    TxtCode.Text = Book.Next_Code(rs_bk, CmbType.Text) 'GENERATE NEXT CODE
    
    Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
    CmdSave.Enabled = True
    CmdExit.Caption = "&Cancel"
    TxtTitle.SetFocus
End Sub

Private Sub CmdDel_Click()
    Dim rs_isu As New Recordset
    Dim qty As Integer, qtyStr As String
    
    'NO DELETE BOOK WHEN IT IS ISSUED
    Set rs_isu = New Recordset
    rs_isu.Open "SELECT * FROM Issue_Mast WHERE [Bk_No]='" & TxtCode & "'", conn, adOpenKeyset
    

    If MsgBox("You want to delete this record ?", vbInformation + vbYesNo, "Book deletion") = vbYes Then
    
        qtyStr = InputBox("Enter Book/CD quantity to delete.", "Book/CD deletion")
                    
        If Trim(qtyStr) = "" Or (Not IsNumeric(qtyStr)) Then
            MsgBox "No proper argument.", vbExclamation, "Book/CD deletion"
            Exit Sub
        Else
            qty = Val(qtyStr)
        End If
                    
        'WHEN ENTERED QUANTITY IS GREATER THAN AVAILABLE QUANTITY
        If qty > TxtAvlQty Then
            MsgBox "Quantity of book/CD you entered is greater than available quantity.", _
                    vbInformation, "Book/CD deletion"
            Exit Sub
        End If
        
        'WHEN BOOK/CD NOT IN LIBRARY
        If qty = TxtQty Then
            If MsgBox("Quantity of book/CD you entered to delete is equal to total quantity." & _
                vbCrLf & "You want to deleted record ?", vbQuestion + vbYesNo, "Book/CD deletion") = vbYes Then
                'code to delete record
                rs_bk.Delete    'DELETE RECORD
                rs_bk.Update    'UPDATE RECORD
                rs_bk.MoveNext  'MOVE RECORDSET TO NEXT RECORD
            End If
        Else
            rs_bk.Fields(6) = rs_bk.Fields(6) - qty
            rs_bk.Update
        End If

        
        'RETRIVE UPDATED DATA
        Call Book.fillGrid(Me, srchCategory, rs_bk)
        If rs_bk.RecordCount = 0 Then
            Call Book.clearText(Me)  'CLEAR TEXT BOXES
            Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
            CmdAdd.Enabled = True   'ENABLE ADD COMMAND BUTTONS
            Exit Sub
        Else
            If rs_bk.EOF Then
                rs_bk.MoveFirst
                Call Book.bookData(Me, rs_bk) 'RETRIVE RECORD
                Exit Sub
            Else
                Call Book.bookData(Me, rs_bk) 'RETRIVE RECORD
            End If
        End If
    End If
End Sub

Private Sub CmdEdit_Click()
    cmd = "Edit"
    Call Book.enableText(Me) 'ENABLE TEXT BOXES
    TxtCode.Locked = True 'DISABLE CODE TEXT BOXES
    CmbType.Enabled = False 'DISADBLE TYPE SELECTION
    Call Book.disableCommand(Me) 'DISABLE COMMAND BUTTONS
    CmdSave.Enabled = True
    CmdExit.Caption = "&Cancel"
End Sub

Private Sub CmdExit_Click()
    If CmdExit.Caption = "&Cancel" Then
        Call Book.clearText(Me)  'CLEAR TEXT BOXEX
        If rs_bk.RecordCount <> 0 Then
            rs_bk.MoveFirst
            Call Book.enableCommand(Me) 'ENABLE COMMAND BUTTON
            Call Book.bookData(Me, rs_bk) 'RETRIVE RECORD
        Else
            Call Book.disableCommand(Me)
            CmdAdd.Enabled = True
        End If
        Call Book.lockText(Me)   'LOCK TEXT BOXES
        CmdSave.Enabled = False
        CmbType.Enabled = True 'ENABLE TYPE SELECTION
        CmdExit.Caption = "E&xit"
    ElseIf CmdExit.Caption = "E&xit" Then
        CmdExit.Caption = "&Cancel"
        Unload Me
    End If
End Sub

Private Sub CmdFirst_Click()
    rs_bk.MoveFirst 'MOVE RECORD TO FIRST
    Call Book.bookData(Me, rs_bk)    'RETRIVE DATA
End Sub

Private Sub CmdLast_Click()
    rs_bk.MoveLast
    Call Book.bookData(Me, rs_bk)    'RETRIVE DATA
End Sub

Private Sub CmdNext_Click()
    rs_bk.MoveNext
    If rs_bk.EOF Then
        rs_bk.MoveFirst
    End If
    Call Book.bookData(Me, rs_bk)    'RETRIVE DATA
End Sub

Private Sub CmdPrv_Click()
    rs_bk.MovePrevious
    If rs_bk.BOF Then
        rs_bk.MoveLast
    End If
    Call Book.bookData(Me, rs_bk)    'RETRIVE DATA
End Sub

Private Sub CmdSave_Click()
    Dim qAdd As String, qEdit As String, dt As String
    
    'CODE VALIDATION
    If Book.codeValid(TxtCode.Text, "Adding New Book", bkType) Then
        TxtCode.SetFocus
        Exit Sub
    End If
    
    'DATA ENTRY VALIDATION
    If Book.dataValid(Me, "Adding New Book") Then
        TxtCode.SetFocus
        Exit Sub
    End If
    
    'CREATE DATE
    dt = CmbDay.Text & "/" & CmbMonth.Text & "/" & CmbYear.Text
    If cmd = "Add" Then
        qAdd = "insert into Book_Mast values('" & TxtCode & _
            "','" & TxtTitle & "','" & TxtAuther & "','" & TxtPub & _
            "','" & dt & "'," & TxtPrice & "," & TxtQty & _
            ",'" & TxtFrom & "'," & TxtAvlQty & ")"
        
        'CHECK FOR DUPLICATION
        If rs_bk.RecordCount <> 0 Then
            rs_bk.Find "Code = '" & TxtCode & "'"
            If Not rs_bk.EOF Then
                MsgBox "Code is already exist.", vbInformation, "Adding New Book"
                Exit Sub
            End If
        End If
        
        conn.Execute qAdd
        Call CmbType_Click  'CALL CLICK METHOD OF TYPE SELECTION COMBOBOX TO RETRIVE REFRESH DATA
        
        Call CmdExit_Click  'CALLING CLICK METHOD OF CMDEXIT
    
    'CODING FOR EDITING RECORD
    ElseIf cmd = "Edit" Then
        qEdit = "update Book_Mast set Title='" & TxtTitle & _
            "', Author='" & TxtAuther & _
            "', Pur_Dt='" & dt & "', Price=" & TxtPrice & _
            ", Qty=" & TxtQty & _
            ", Pur_From='" & TxtFrom & _
            "' where code='" & TxtCode & "'"
            
        MsgBox qEdit
        conn.Execute qEdit
        Call Book.fillGrid(Me, "Code", rs_bk) 'FILL FLEX GRID
        Call CmdExit_Click  'CALLING CLICK METHOD OF CMDEXIT
    End If
End Sub

Private Sub CmdSrch_Click()
    Unload Me
    FrmSearch.Show vbModal
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'FILL COMBO OF DATES
    For i = 1 To DateDiff("d", Date, DateAdd("m", 1, Date)) 'fill days
        CmbDay.AddItem i
    Next
    For i = 1 To 12 'fill months
        CmbMonth.AddItem i
    Next
    For i = 1950 To 2050 'fill years
        CmbYear.AddItem i
    Next
    
    CmbType.Text = bkType ''SET BOOK DATA AS DEFAULT TYPE
    'LOCK TEXT BOXES
    Call Book.lockText(Me)
    
    Me.MsfgSearch.FormatString = "No. |Code    |Title                    |Auther                  |Publisher                " & _
                                "|Purchase Date |Price   |Quantity|Available Quantity"
    

End Sub


Private Sub Form_Resize()
    'SET LABEL
    ShapLabel.Width = Me.ScaleWidth
    LblLabel.Left = ShapLabel.Width / 2 - LblLabel.Width / 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs_bk.Close
    
    If Forms.Count = 2 Then
        MDIFrm.Pct1.Visible = True
    End If
End Sub

Private Sub MsfgSearch_Click()
    rs_bk.MoveFirst
    rs_bk.Find "Code = '" & MsfgSearch.TextMatrix(MsfgSearch.Row, 1) & "'"
    
    Call Book.bookData(Me, rs_bk)
End Sub

Private Sub MsfgSearch_SelChange()
    rs_bk.MoveFirst
    rs_bk.Find "Code = '" & MsfgSearch.TextMatrix(MsfgSearch.Row, 1) & "'"
    
    Call Book.bookData(Me, rs_bk)
End Sub

Private Sub OptAuthor_Click()
    TxtSearch.Text = ""
    srchCategory = "Author"
    OptAuthor.Value = True
    Call Book.fillGrid(Me, srchCategory, rs_bk) 'FILL FLEX GRID
End Sub

Private Sub OptCode_Click()
    TxtSearch.Text = ""
    srchCategory = "Code"
    OptCode.Value = True
    Call Book.fillGrid(Me, srchCategory, rs_bk) 'FILL FLEX GRID
End Sub

Private Sub OptPublisher_Click()
    TxtSearch.Text = ""
    srchCategory = "Publisher"
    OptPublisher.Value = True
    Call Book.fillGrid(Me, srchCategory, rs_bk) 'FILL FLEX GRID
End Sub

Private Sub OptTitle_Click()
    TxtSearch.Text = ""
    srchCategory = "Title"
    OptTitle.Value = True
    Call Book.fillGrid(Me, srchCategory, rs_bk) 'FILL FLEX GRID
End Sub

Private Sub TxtAuther_GotFocus()
    Call Book.selectTxt(TxtAuther)
End Sub

Private Sub TxtAuther_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtCode_GotFocus()
    Call Book.selectTxt(TxtCode)
End Sub

Private Sub TxtCode_KeyPress(KeyAscii As Integer)
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtFrom_GotFocus()
    Call Book.selectTxt(TxtFrom)
End Sub

Private Sub TxtFrom_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtPrice_GotFocus()
    Call Book.selectTxt(TxtPrice)
End Sub

Private Sub TxtPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = 8
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPub_GotFocus()
    Call Book.selectTxt(TxtPub)
End Sub

Private Sub TxtPub_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtQty_GotFocus()
    Call Book.selectTxt(TxtQty)
End Sub

Private Sub TxtQty_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then
        KeyAscii = 8
    ElseIf KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtSearch_Change()
    Set rs_temp = New Recordset
    
    If bkType = "BOOK" Then
        rs_temp.Open "select * from Book_Mast where Code like('B%') and " & srchCategory & " like ('" & TxtSearch & "%') order by " & srchCategory, conn, adOpenStatic, adLockReadOnly
    Else
        rs_temp.Open "select * from Book_Mast where code like('C%') and " & srchCategory & " like ('" & TxtSearch & "%') order by " & srchCategory, conn, adOpenStatic, adLockReadOnly
    End If
         
    If rs_temp.RecordCount = 0 Then
        MsfgSearch.Enabled = False
    Else
        MsfgSearch.Enabled = True
    End If
    
    Call fillGrid1(Me, rs_temp) 'fill grid
    
End Sub

Private Sub TxtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtTitle_GotFocus()
    Call Book.selectTxt(TxtTitle)
End Sub

Private Sub TxtTitle_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub
