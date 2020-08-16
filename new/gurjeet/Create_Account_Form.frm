VERSION 5.00
Begin VB.Form FrmCreateAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create New Account"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7230
   Icon            =   "Create_Account_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7230
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox CmbUserType 
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
      ItemData        =   "Create_Account_Form.frx":030A
      Left            =   3840
      List            =   "Create_Account_Form.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   765
      Width           =   1320
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   5040
      Width           =   1095
   End
   Begin VB.CommandButton CmdCreateAcc 
      Caption         =   "Create &Account"
      Default         =   -1  'True
      Height          =   375
      Left            =   2475
      TabIndex        =   9
      Top             =   4440
      Width           =   2355
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "&Back"
      Height          =   375
      Left            =   840
      TabIndex        =   10
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox TxtConfPwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   20
      PasswordChar    =   "l"
      TabIndex        =   8
      Top             =   3870
      Width           =   5295
   End
   Begin VB.TextBox TxtPwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Wingdings"
         Size            =   9.75
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   960
      MaxLength       =   20
      PasswordChar    =   "l"
      TabIndex        =   6
      Top             =   2790
      Width           =   5295
   End
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   960
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1725
      Width           =   5295
   End
   Begin VB.Label LblUserType 
      AutoSize        =   -1  'True
      Caption         =   "&Select type of user :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1875
      TabIndex        =   1
      Top             =   810
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Type the password to c&onfirm :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   7
      Top             =   3570
      Width           =   3015
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Type a &password :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   2490
      Width           =   1800
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Type your &user name :"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      TabIndex        =   3
      Top             =   1425
      Width           =   2190
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create New Account"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   1755
      TabIndex        =   0
      Top             =   120
      Width           =   3570
   End
   Begin VB.Shape ShapLabel 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "FrmCreateAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_user As New ADODB.Recordset

Private Sub CmdBack_Click()
    Unload Me
    FrmUserMng.Show vbModal
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCreateAcc_Click()
    Dim Query As String, typ As String
    
    If Trim(TxtUser.Text) = "" Or Trim(TxtPwd) = "" Or Trim(TxtConfPwd) = "" Then
        MsgBox "All fields are compulsory.", vbInformation, "User Addition"
        Exit Sub
    ElseIf Trim(TxtPwd) <> Trim(TxtConfPwd) Then
        MsgBox "Your confirm password do not match." & vbCrLf & _
                "Type your confirm password again.", vbCritical, "User Addition"
        TxtConfPwd.SetFocus
        Exit Sub
    End If
    
    'DUPLICATION CHECK
    rs_user.MoveFirst
    rs_user.Find "usr='" & TxtUser & "'"
    
    If rs_user.EOF Then 'USER NOT EXIST
        If CmbUserType.Text = "ADMIN" Then
            typ = "A"
        Else
            typ = "L"
        End If
        
        Query = "insert into Login_Mast values ('" & TxtUser & "','" & _
                    TxtPwd & "','" & typ & "')"
        MsgBox Query
        conn.Execute Query
        MsgBox "New user is successfully added.", vbInformation, "User Addition"
        TxtUser = ""
        TxtPwd = ""
        TxtConfPwd = ""
        CmbUserType.SetFocus
        Call Form_Load
    Else    'USER IS ALREADY EXIST
        MsgBox "User already exit. Enter another user name.", vbCritical, "User Additon"
        TxtUser.SetFocus
    End If
End Sub

Private Sub Form_Load()
    MDIFrm.Pct1.Visible = False
    
    'OPEN RECORDSET
    If rs_user.State = 1 Then rs_user.Close
    rs_user.Open "select * from Login_Mast", conn, adOpenStatic, adLockPessimistic
    
    'CLEAR TEXT BOX
    TxtUser.Text = ""
    TxtPwd.Text = ""
    TxtConfPwd.Text = ""
    
    CmbUserType.Text = CmbUserType.List(0)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs_user.Close
    
End Sub

Private Sub TxtConfPwd_GotFocus()
    Call Book.selectTxt(TxtConfPwd)
End Sub

Private Sub TxtConfPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtPwd_GotFocus()
    Call Book.selectTxt(TxtPwd)
End Sub

Private Sub TxtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtUser_GotFocus()
    Call Book.selectTxt(TxtUser)
End Sub

Private Sub TxtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub
