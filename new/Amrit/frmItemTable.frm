VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmItemTable 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " "
   ClientHeight    =   5310
   ClientLeft      =   6165
   ClientTop       =   3585
   ClientWidth     =   9345
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmItemTable.frx":0000
   ScaleHeight     =   5310
   ScaleWidth      =   9345
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   720
      Top             =   4560
      Visible         =   0   'False
      Width           =   4575
      _ExtentX        =   8070
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Item_Table"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "&Exit"
      Height          =   375
      Left            =   7200
      TabIndex        =   12
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Frame fraNavigator 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Navigation"
      Height          =   1125
      Left            =   720
      TabIndex        =   15
      Top             =   3360
      Width           =   6375
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Last"
         Height          =   465
         Index           =   3
         Left            =   4800
         TabIndex        =   11
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Next"
         Height          =   465
         Index           =   2
         Left            =   3250
         TabIndex        =   10
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "Previous"
         Height          =   465
         Index           =   1
         Left            =   1700
         TabIndex        =   9
         Top             =   450
         Width           =   1455
      End
      Begin VB.CommandButton cmdNavigate 
         Caption         =   "First"
         Height          =   465
         Index           =   0
         Left            =   150
         TabIndex        =   8
         Top             =   450
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3520
      Left            =   7080
      TabIndex        =   14
      Top             =   960
      Width           =   1815
      Begin VB.CommandButton cmdmodify 
         Caption         =   "&Modify"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdupdate 
         Caption         =   "&Update"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   2400
         Width           =   1575
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   3000
         Width           =   1575
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "&Save"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   810
         Width           =   1575
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   1845
         Width           =   1575
      End
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add New"
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   330
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2385
      Left            =   720
      TabIndex        =   13
      Top             =   960
      Width           =   6375
      Begin VB.TextBox Text1 
         DataField       =   "Cat_id"
         DataSource      =   "Adodc1"
         Height          =   390
         Left            =   5520
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2340
         TabIndex        =   1
         Top             =   300
         Width           =   3135
      End
      Begin VB.TextBox txtInput 
         DataField       =   "Rate"
         DataSource      =   "Adodc1"
         Height          =   390
         Index           =   1
         Left            =   2310
         MaxLength       =   6
         TabIndex        =   3
         Top             =   1440
         Width           =   3135
      End
      Begin VB.TextBox txtInput 
         DataField       =   "Item_name"
         DataSource      =   "Adodc1"
         Height          =   390
         Index           =   0
         Left            =   2325
         TabIndex        =   2
         Top             =   870
         Width           =   3135
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   615
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   1530
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name"
         Height          =   615
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Category Name"
         Height          =   615
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Item Information"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1920
      TabIndex        =   19
      Top             =   240
      Width           =   5715
   End
End
Attribute VB_Name = "frmItemTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function EDControls(mode As Boolean)
    txtInput(0).Enabled = mode
    txtInput(1).Enabled = mode
    Combo1.Enabled = mode
    cmdAddNew.Enabled = Not mode
    cmdCancel.Enabled = mode
    cmdFind.Enabled = Not mode
    cmdDelete.Enabled = Not mode
    cmdmodify.Enabled = Not mode
End Function

Private Function EDNavigate(mode As Boolean)
    fraNavigator.Enabled = mode
End Function

Private Sub cmdAddNew_Click()
    EDControls True
    EDNavigate False
    cmdSave.Enabled = True
    cmdCancel.Enabled = True
    Adodc1.Recordset.AddNew
End Sub

Private Sub cmdCancel_Click()
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    Adodc1.Refresh
    Adodc1.Recordset.MoveFirst
End Sub

Private Sub cmdDelete_Click()
    Dim choice As Integer
        choice = MsgBox("Do you want to Delete the Record", vbYesNo + vbQuestion, "confirmation")
        If choice = vbYes Then
            If Adodc1.Recordset.EOF = False And Adodc1.Recordset.BOF = False Then
                Adodc1.Recordset.Delete
                Adodc1.Recordset.MoveNext
                If Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveLast
            End If
        End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
Dim s As String
    If Adodc1.Recordset.RecordCount > 0 Then
        Adodc1.Recordset.MoveFirst
        s = InputBox("Enter The Item Name")
        If s <> "" Then
            Adodc1.Recordset.Find "item_name='" & s & "'"
            If Not Adodc1.Recordset.EOF Then
            Else
                MsgBox "Item Not Found"
                Adodc1.Recordset.MoveLast
            End If
        End If
        
    End If
End Sub

Private Sub cmdModify_Click()
    EDControls True
    EDNavigate False
    cmdSave.Enabled = False
    cmdupdate.Enabled = True
    txtInput(0).SetFocus
    cmdCancel.Enabled = False
    cmdmodify.Visible = False
End Sub

Private Sub cmdNavigate_click(Index As Integer)
      Select Case Index
      Case 0
            Adodc1.Recordset.MovePrevious
            If (Adodc1.Recordset.BOF = True) Then
                MsgBox "You are already at the First Record"
            End If
            Adodc1.Recordset.MoveFirst
      Case 1
            Adodc1.Recordset.MovePrevious
            If Adodc1.Recordset.BOF Then
                MsgBox "You are already at the First Record"
                Adodc1.Recordset.MoveFirst
            End If
      Case 2
            Adodc1.Recordset.MoveNext
            If Adodc1.Recordset.EOF Then
            MsgBox "you are already at the last record"
            Adodc1.Recordset.MoveLast
            End If
      Case 3
            Adodc1.Recordset.MoveNext
            If (Adodc1.Recordset.EOF = True) Then
            MsgBox "You are already at the last Record"
            End If
            Adodc1.Recordset.MoveLast
      End Select
    Combo1.ListIndex = CInt(Text1.Text) - 1
End Sub

Private Sub cmdSave_Click()
    If txtInput(0).Text = "" Or txtInput(1).Text = "" Then
        MsgBox "please provide the data"
        Exit Sub
    End If
    EDControls False
    EDNavigate True
    cmdSave.Enabled = False
    Adodc1.Recordset.Fields("Item_name") = txtInput(0).Text
    Adodc1.Recordset.Fields("Rate") = txtInput(1).Text
    Adodc1.Recordset.Fields("Cat_id") = Combo1.ListIndex + 1
    Adodc1.Recordset.Update
End Sub

Private Sub cmdUpdate_Click()
    If txtInput(0).Text = "" Or txtInput(1).Text = "" Then
        MsgBox "Please provide the data"
        Exit Sub
    End If
    EDControls False
    EDNavigate True
    cmdmodify.Visible = True
    cmdupdate.Enabled = False
End Sub

Private Sub Form_Load()
    Call loadCategory
    Adodc1.RecordSource = "Select * from Item_Table"
    Adodc1.Refresh
    EDControls False
    cmdSave.Enabled = False
    cmdupdate.Enabled = False
    Adodc1.Recordset.MoveFirst
    Combo1.ListIndex = CInt(Text1.Text) - 1
End Sub


Private Function loadCategory()
    Adodc1.RecordSource = "select * from Category"
    Adodc1.Refresh
    Combo1.Clear
    If Adodc1.Recordset.RecordCount > 0 Then
        While Not Adodc1.Recordset.EOF
            Combo1.AddItem Adodc1.Recordset.Fields("cat_name")
            Adodc1.Recordset.MoveNext
        Wend
    End If
End Function

Private Sub txtInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
         If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 46)) Then
            KeyAscii = 0
            MsgBox "Please Enter Numeric Value "
        End If
    End If
End Sub
