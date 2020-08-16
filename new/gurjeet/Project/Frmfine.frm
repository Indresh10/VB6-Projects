VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Frmfine 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Form1"
   ClientHeight    =   3240
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   5265
   StartUpPosition =   1  'CenterOwner
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   120
      Top             =   2760
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
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
      Connect         =   $"Frmfine.frx":0000
      OLEDBString     =   $"Frmfine.frx":008C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select * from Fine where Status='Not Paid'"
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
   Begin VB.CommandButton Command1 
      Caption         =   "Payment"
      Height          =   495
      Left            =   3240
      TabIndex        =   9
      Top             =   2640
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Fee Payment"
      Height          =   2415
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         Top             =   1800
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2160
         TabIndex        =   6
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         Top             =   840
         Width           =   2295
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fine Amount"
         Height          =   195
         Left            =   720
         TabIndex        =   7
         Top             =   1830
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Fine Date"
         Height          =   195
         Left            =   720
         TabIndex        =   5
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Book Code"
         Height          =   195
         Left            =   720
         TabIndex        =   3
         Top             =   870
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0FFC0&
         Caption         =   "Member code"
         Height          =   195
         Left            =   720
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "Frmfine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Combo1_Click()
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("MCode='" + Combo1.Text + "'")
Text1.Text = Adodc1.Recordset.Fields(1)
Text2.Text = Adodc1.Recordset.Fields(3)
Text3.Text = Adodc1.Recordset.Fields(2)
End Sub

Private Sub Command1_Click()
If Combo1.Text = "" Then
    MsgBox "please select the member code", vbExclamation
Else
Adodc1.Recordset.Update "Status", "Paid"
Adodc1.Recordset.Update 3, Format(Date, "dd-MMM-yyyy")
Adodc1.Refresh
Adodc1.RecordSource = "Select * from Member where Code='" + Combo1.Text + "'"
Adodc1.Refresh
Adodc1.Recordset.Update "Fine", 0
'MDIFrm.Txtsearch2.Text = "any"
MDIFrm.Txtsearch2.Text = ""
Unload Me
End If
End Sub

Private Sub Form_Load()
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then Adodc1.Recordset.MoveFirst
While Not Adodc1.Recordset.EOF
    Combo1.AddItem Adodc1.Recordset.Fields(0)
    Adodc1.Recordset.MoveNext
Wend
End Sub
