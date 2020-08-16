VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmChangePassword 
   BackColor       =   &H8000000D&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Change Password"
   ClientHeight    =   3120
   ClientLeft      =   7860
   ClientTop       =   4440
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5040
   Begin VB.TextBox txtConfirmPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1650
      Width           =   2310
   End
   Begin VB.TextBox txtNewPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   1140
      Width           =   2295
   End
   Begin VB.TextBox txtCurrentPassword 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2520
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   600
      Width           =   2310
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "CANCEL"
      Height          =   705
      Left            =   2760
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   750
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "SAVE"
      Default         =   -1  'True
      Height          =   705
      Left            =   840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   750
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   1440
      Top             =   2400
      Width           =   1455
      _ExtentX        =   2566
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password : -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   2190
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "New Password : -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   6
      Top             =   1140
      Width           =   1830
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Old Password : -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   270
      Left            =   240
      TabIndex        =   5
      Top             =   600
      Width           =   1710
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H00404040&
      Height          =   1875
      Left            =   120
      Top             =   240
      Width           =   4755
   End
End
Attribute VB_Name = "frmChangePassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub
Private Sub cmdSave_Click()
    If txtCurrentPassword.Text = "" Then
        MsgBox "Please Enter Old Password ...", vbExclamation
        txtCurrentPassword.SetFocus
        Exit Sub
    End If
    
    If txtConfirmPassword.Text = "" Then
    MsgBox "Enter confirm password ...", vbExclamation
    txtConfirmPassword.SetFocus
    Exit Sub
    End If
    
    If txtNewPassword.Text <> txtConfirmPassword.Text Then
        MsgBox "Confirm password does not match with new password ...", vbExclamation
        txtConfirmPassword.Text = ""
        txtNewPassword.Text = ""
        txtNewPassword.SetFocus
        Exit Sub
    End If
Adodc1.RecordSource = "select * from Login where User_name='" + Text1.Text + "'"
Adodc1.Refresh
If Not Adodc1.Recordset.EOF Then
    Adodc1.Recordset.Update 1, txtNewPassword.Text
    MsgBox "password changed successfully"
    Unload Me
Else
    MsgBox "please check your password"
    txtConfirmPassword.Text = ""
    txtNewPassword.Text = ""
    txtCurrentPassword.Text = ""
    txtCurrentPassword.SetFocus
End If
End Sub
