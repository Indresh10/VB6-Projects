VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmset 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1080
      Top             =   3360
      Visible         =   0   'False
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   582
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"frmset.frx":0000
      OLEDBString     =   $"frmset.frx":008C
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Settings"
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
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   495
      Left            =   3120
      TabIndex        =   6
      Top             =   2880
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Save"
      Height          =   495
      Left            =   840
      TabIndex        =   5
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Height          =   1815
      Left            =   615
      TabIndex        =   0
      Top             =   720
      Width           =   4335
      Begin VB.TextBox Text2 
         Height          =   400
         Left            =   2400
         TabIndex        =   4
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Height          =   400
         Left            =   2400
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.Image Image2 
         Height          =   465
         Left            =   2040
         Picture         =   "frmset.frx":0118
         Top             =   950
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   465
         Left            =   2040
         Picture         =   "frmset.frx":0589
         Top             =   345
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Member Fees"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fees Amount per Day"
         Height          =   195
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1530
      End
   End
End
Attribute VB_Name = "frmset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Fields(0) = Val(Text1.Text)
Adodc1.Recordset.Fields(1) = Val(Text2.Text)
Adodc1.Recordset.Update
Adodc1.Refresh
MsgBox "Settings Saved successfully", vbInformation
MDIFrm.Label12.Caption = Text1.Text
MDIFrm.Label15.Caption = Text2.Text
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
Text1.Text = Adodc1.Recordset.Fields(0)
Text2.Text = Adodc1.Recordset.Fields(1)
Adodc1.Refresh
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 46)) Then
            KeyAscii = 0
            MsgBox "Please Enter Numeric Value "
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 46)) Then
            KeyAscii = 0
            MsgBox "Please Enter Numeric Value "
End If
End Sub

