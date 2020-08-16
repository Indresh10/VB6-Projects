VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form FrmEditAcc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Account"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   Icon            =   "Edit_User_Account_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   6975
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   5010
      TabIndex        =   20
      Top             =   6000
      Width           =   1095
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "&Back"
      Height          =   375
      Left            =   765
      TabIndex        =   19
      Top             =   6000
      Width           =   1095
   End
   Begin TabDlg.SSTab TabEditAcc 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   495
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Change &User Name"
      TabPicture(0)   =   "Edit_User_Account_Form.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FremEditUser"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Change &Password"
      TabPicture(1)   =   "Edit_User_Account_Form.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FremEditPwd"
      Tab(1).ControlCount=   1
      Begin VB.Frame FremEditPwd 
         Height          =   4815
         Left            =   -74880
         TabIndex        =   9
         Top             =   360
         Width           =   6735
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
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   20
            PasswordChar    =   "l"
            TabIndex        =   17
            Top             =   3525
            Width           =   5295
         End
         Begin VB.TextBox TxtNewPwd 
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
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   20
            PasswordChar    =   "l"
            TabIndex        =   15
            Top             =   2565
            Width           =   5295
         End
         Begin VB.CommandButton CmdPwd 
            Caption         =   "C&hange password"
            Height          =   375
            Left            =   2040
            TabIndex        =   18
            Top             =   4320
            Width           =   2655
         End
         Begin VB.TextBox TxtCurPwd 
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
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   720
            MaxLength       =   20
            PasswordChar    =   "l"
            TabIndex        =   13
            Top             =   1605
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
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   11
            Top             =   645
            Width           =   5295
         End
         Begin VB.Label LblConfPwd 
            AutoSize        =   -1  'True
            Caption         =   "C&onfirm password :"
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
            Left            =   720
            TabIndex        =   16
            Top             =   3240
            Width           =   1920
         End
         Begin VB.Label LblNewPwd 
            AutoSize        =   -1  'True
            Caption         =   "&Type a new password :"
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
            Left            =   720
            TabIndex        =   14
            Top             =   2280
            Width           =   2265
         End
         Begin VB.Label LblCurPwd 
            AutoSize        =   -1  'True
            Caption         =   "Enter your current pa&ssword"
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
            Left            =   720
            TabIndex        =   12
            Top             =   1320
            Width           =   2775
         End
         Begin VB.Label LblUser 
            AutoSize        =   -1  'True
            Caption         =   "Enter user &name :"
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
            Left            =   720
            TabIndex        =   10
            Top             =   360
            Width           =   1755
         End
      End
      Begin VB.Frame FremEditUser 
         Height          =   4815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   6735
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
            ItemData        =   "Edit_User_Account_Form.frx":0342
            Left            =   3540
            List            =   "Edit_User_Account_Form.frx":034C
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   570
            Width           =   1320
         End
         Begin VB.CommandButton CmdUser 
            Caption         =   "C&hange user name"
            Height          =   375
            Left            =   2040
            TabIndex        =   8
            Top             =   4320
            Width           =   2655
         End
         Begin VB.TextBox TxtNewUser 
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
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   7
            Top             =   3060
            Width           =   5295
         End
         Begin VB.TextBox TxtCurUser 
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
            Height          =   285
            Left            =   720
            MaxLength       =   20
            TabIndex        =   5
            Top             =   1785
            Width           =   5295
         End
         Begin VB.Label LblUserType 
            AutoSize        =   -1  'True
            Caption         =   "Change user &type :"
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
            Left            =   1575
            TabIndex        =   2
            Top             =   615
            Width           =   1845
         End
         Begin VB.Label LblNewUser 
            AutoSize        =   -1  'True
            Caption         =   "Type a &new user name :"
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
            Left            =   720
            TabIndex        =   6
            Top             =   2700
            Width           =   2355
         End
         Begin VB.Label LblCurUser 
            AutoSize        =   -1  'True
            Caption         =   "Cu&rrent user name :"
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
            Left            =   720
            TabIndex        =   4
            Top             =   1425
            Width           =   1965
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Edit existing user account"
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
      Left            =   1200
      TabIndex        =   21
      Top             =   15
      Width           =   4485
   End
   Begin VB.Shape ShapLabel 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   6975
   End
End
Attribute VB_Name = "FrmEditAcc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_user As New ADODB.Recordset
Dim Query As String

Private Sub CmdBack_Click()
    Unload Me
    FrmUserMng.Show vbModal
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdPwd_Click()
    'CHECKING FOR BLANCK TEXT BOXES
    If Trim(TxtUser) = "" Or Trim(TxtCurPwd) = "" Or Trim(TxtNewPwd) = "" Or Trim(TxtConfPwd) = "" Then
        MsgBox "All fields are compulsory.", vbInformation, "Change Password"
        Exit Sub
    End If
    'CHECKING FOR NEW PASSWORD & CONF. PASSWORD MATCHING
    If TxtNewPwd <> TxtConfPwd Then
        MsgBox "Your confirm password do not match." & vbCrLf & "Enter confirm password again.", vbCritical, "Change Password"
        Exit Sub
    End If
    
    If userType = "A" Then
        'WHEN USER IS ADMIN
        'FIND USER IS EXIST OR NOT
        rs_user.MoveFirst
        rs_user.Find "usr='" & TxtUser & "'"
    
        If rs_user.EOF Then 'USER NOT EXIST
            MsgBox "User name does not exixt." & vbCrLf & _
                "Enter current name again.", vbCritical, "Change Password"
            TxtUser.SetFocus
            Exit Sub
        End If
        
        'IF USER AND PASSWORD NOT MATCH
        If rs_user.Fields(1) <> TxtCurPwd Then
            MsgBox "Your current password do not match. Enter it again.", vbCritical, "Change Password"
            TxtCurPwd.SetFocus
            Exit Sub
        End If
        
        Query = "update Login_Mast set pw='" & TxtNewPwd & "' where usr='" & TxtUser & "'"
    
    Else
        'WHEN USER IS LIMITED
        If TxtUser <> userNm Then
            MsgBox "Your user name not match. Enter it again.", vbCritical, "Change Password"
            Exit Sub
        End If
        
        rs_user.MoveFirst
        rs_user.Find "usr='" & TxtUser & "'"
        
        'IF USER AND PASSWORD NOT MATCH
        If rs_user.Fields(1) <> TxtCurPwd Then
            MsgBox "Your current password do not match. Enter it again.", vbCritical, "Change Password"
            TxtCurPwd.SetFocus
            Exit Sub
        End If
        
        Query = "update Login_Mast set pw='" & TxtNewPwd & "' where usr='" & TxtUser & "'"
    
    End If
    
    'UPDATE PASSWORD
    conn.Execute Query
    MsgBox "Your password is changed successfully.", vbInformation, "Change Password"
    
    TxtUser.Text = ""
    TxtCurPwd.Text = ""
    TxtNewPwd.Text = ""
    TxtConfPwd.Text = ""
    TxtUser.SetFocus
End Sub

Private Sub CmdUser_Click()
    Dim typ As String
    typ = userType
    
    If Trim(TxtCurUser) = "" Or Trim(TxtNewUser) = "" Then
        MsgBox "All fields are compulsory.", vbInformation, "Change User"
        Exit Sub
    End If
    
    If userType = "A" Then
        'WHEN USER IS ADMIN
        'FIND USER IS EXIST OR NOT
        rs_user.MoveFirst
        rs_user.Find "usr='" & TxtCurUser & "'"
    
        If rs_user.EOF Then 'USER NOT EXIST
            MsgBox "User name does not exixt." & vbCrLf & _
                "Enter current name again.", vbCritical, "User Edition"
            TxtCurUser.SetFocus
            Exit Sub
        End If
          
        If CmbUserType.Text = "ADMIN" Then
            typ = "A"
        Else
            typ = "L"
        End If
        
        'WHEN CURRENT USER IS CHANGING ACCOUNT
        If (TxtCurUser = userNm) And (userType <> typ) Then
            MsgBox "You can not change your account type." & vbCrLf & _
                "Login with another Admin user and then change your account type.", vbInformation, "User Edition"
            Exit Sub
        End If
        
        Query = "update Login_Mast set usr='" & TxtNewUser & "',typ='" & _
                typ & "' where usr='" & TxtCurUser & "'"
                
    Else
        
        'WHEN USER IS LIMITED
        If userNm <> TxtCurUser Then
            MsgBox "Your current name is not correct." & vbCrLf & _
                "Enter current name again.", vbCritical, "User Edition"
            TxtCurUser.SetFocus
            Exit Sub
        End If
        
        Query = "update Login_Mast set usr='" & TxtNewUser & "' where usr='" & userNm & "'"
    End If
    
    'CHECK FOR DUPLICATE RECORD
    rs_user.MoveFirst
    rs_user.Find "usr='" & TxtNewUser & "'"
    
    If (rs_user.EOF = False) And (userNm <> TxtNewUser) Then
        MsgBox "User already exixt. Enter another user name.", vbCritical, "User Edition"
        Exit Sub
    End If
    
    'EXECUTE QUERY & UPDATE RECORD
    conn.Execute Query
    MsgBox "User name is changed successfully.", vbInformation, "User Edition"
    
    If TxtCurUser = userNm Then
        userNm = TxtNewUser.Text
        userType = typ
        MDIFrm.StatusBar1.Panels(1) = "Current User : " & userNm & "(" & userType & ")"
    End If
    
    TxtCurUser.Text = ""
    TxtNewUser.Text = ""
    TxtCurUser.SetFocus
End Sub

Private Sub Form_Load()
    MDIFrm.Pct1.Visible = False
    
    Call TabEditAcc_Click(0)    'SELECT TAB 1
    
    'OPEN RECORDSET
    If rs_user.State = 1 Then rs_user.Close 'CLOSE RECORDSET IF OPEN
    rs_user.Open "select * from Login_Mast", conn, adOpenStatic, adLockPessimistic
    
End Sub

Private Sub TabEditAcc_Click(PreviousTab As Integer)
    If TabEditAcc.Tab = 0 Then
        FremEditPwd.Enabled = False
        FremEditUser.Enabled = True
        CmdUser.Default = True
        
        'CmbUserType.Text = CmbUserType.List(0)
        TxtCurUser = ""
        TxtNewUser = ""
        
        If userType = "L" Then
            CmbUserType.Enabled = False
            CmbUserType.Text = "LIMITED"
        Else
            CmbUserType.Text = "ADMIN"
        End If
        
    Else
        FremEditPwd.Enabled = True
        FremEditUser.Enabled = False
        CmdPwd.Default = True
        
        TxtUser = "":   TxtCurPwd = ""
        TxtNewPwd = "":     TxtConfPwd = ""
    End If
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

Private Sub TxtCurPwd_GotFocus()
    Call Book.selectTxt(TxtCurPwd)
End Sub

Private Sub TxtCurPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtCurUser_GotFocus()
    Call Book.selectTxt(TxtCurUser)
End Sub

Private Sub TxtCurUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtNewPwd_GotFocus()
    Call Book.selectTxt(TxtNewPwd)
End Sub

Private Sub TxtNewPwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 34 Or KeyAscii = 32 Then
        KeyAscii = 0
    End If
    KeyAscii = Book.upper(KeyAscii)
End Sub

Private Sub TxtNewUser_GotFocus()
    Call Book.selectTxt(TxtNewUser)
End Sub

Private Sub TxtNewUser_KeyPress(KeyAscii As Integer)
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
