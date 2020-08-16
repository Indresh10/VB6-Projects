VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAddNewUser 
   BackColor       =   &H000000FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add New User"
   ClientHeight    =   3105
   ClientLeft      =   7860
   ClientTop       =   4230
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   4935
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1800
      Top             =   2520
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
   Begin VB.CommandButton cmdClose 
      Caption         =   "CANCEL"
      Height          =   705
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   750
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Default         =   -1  'True
      Height          =   705
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   750
   End
   Begin VB.TextBox txtPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   2175
   End
   Begin VB.TextBox txtConfirmPassword 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2175
   End
   Begin VB.TextBox txtUsername 
      Appearance      =   0  'Flat
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2520
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblPassword4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label lblConfirmPassword4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password :-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblUsername4 
      BackColor       =   &H00C0C000&
      BackStyle       =   0  'Transparent
      Caption         =   "Username:-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   480
      Width           =   2055
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00404040&
      Height          =   1875
      Left            =   120
      Top             =   240
      Width           =   4635
   End
End
Attribute VB_Name = "frmAddNewUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
    If txtUsername.Text = "" Or txtPassword.Text = "" Then
        MsgBox "Enter UserName and Password ...", vbExclamation
        txtUsername.SetFocus
        Exit Sub
    End If
    If txtPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "Confirm password dosenot match with new password ...", vbExclamation
        txtConfirmPassword.Text = ""
        txtPassword.Text = ""
        txtPassword.SetFocus
        Exit Sub
    End If
Adodc1.RecordSource = ("select * from Login where User_name='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    MsgBox "Sorry!! User already exists. Try another username", vbCritical
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
Else
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields(0) = txtUsername.Text
    Adodc1.Recordset.Fields(1) = txtPassword.Text
    Adodc1.Recordset.Update
    MsgBox "User added sucessfully", vbInformation
    txtPassword.Text = ""
    txtConfirmPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
   Unload Me
End If
End Sub

