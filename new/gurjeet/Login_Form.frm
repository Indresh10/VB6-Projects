VERSION 5.00
Begin VB.Form FrmLogin 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5775
   Icon            =   "Login_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   5775
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox TxtUser 
      Appearance      =   0  'Flat
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
      Left            =   2160
      MaxLength       =   20
      TabIndex        =   1
      Top             =   600
      Width           =   2655
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
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   2160
      MaxLength       =   20
      PasswordChar    =   "l"
      TabIndex        =   3
      Top             =   1440
      Width           =   2655
   End
   Begin VB.CommandButton CmdLogin 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   840
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Label LblUser 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name :"
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
      Left            =   840
      TabIndex        =   0
      Top             =   660
      Width           =   1125
   End
   Begin VB.Label LblPwd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "&Password :"
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
      Left            =   840
      TabIndex        =   2
      Top             =   1485
      Width           =   990
   End
   Begin VB.Image ImgLogin 
      Height          =   3120
      Left            =   0
      Picture         =   "Login_Form.frx":29D5A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5865
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_user As New Recordset

Private Sub CmdCancel_Click()
    End
End Sub

Private Sub CmdLogin_Click()
    bkType = "BOOK"
    userType = "L"
    Class = "BBA": Yer = "FY"
    
    If Trim(TxtUser) = "" And Trim(TxtPwd) = "" Then
        MsgBox "Fill all the details", vbInformation, "Login"
        TxtUser.SetFocus
        Exit Sub
    End If
    
    If TxtUser = "LIBRARY" And TxtPwd = "INDISOFT" Then
        userType = "L"
        userNm = "LIBRARY"
        Unload FrmWelcome
        Unload FrmLogin
        MDIFrm.Show
        Exit Sub
    End If
    
    If rs_user.RecordCount <> 0 Then
        rs_user.MoveFirst
        rs_user.Find "usr = '" & TxtUser & "'"
        If Not rs_user.EOF Then
            If rs_user.Fields(1) = TxtPwd Then
                userType = rs_user.Fields(2)
                userNm = rs_user.Fields(0)
                Unload FrmWelcome
                Unload FrmLogin
                MDIFrm.Show
                Exit Sub
            Else
                MsgBox "Wrong username or password.", vbCritical, "Login"
                TxtUser.SetFocus
                Exit Sub
            End If
        Else
            MsgBox "Wrong username or password.", vbCritical, "Login"
            TxtUser.SetFocus
            Exit Sub
        End If
    Else
        MsgBox "Wrong username or password.", vbCritical, "Login"
        TxtUser.SetFocus
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    'OPEN RECORDSET
    rs_user.Open "select * from Login_Mast", conn, adOpenStatic, adLockPessimistic
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs_user.Close
End Sub

Private Sub TxtPwd_GotFocus()
    Call selectTxt(TxtPwd)
End Sub

Private Sub TxtPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = upper(KeyAscii)
End Sub

Private Sub TxtUser_GotFocus()
    Call selectTxt(TxtUser)
End Sub

Private Sub TxtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then
        KeyAscii = 0
    End If
    KeyAscii = upper(KeyAscii)
End Sub
