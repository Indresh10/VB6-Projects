VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmdata_man 
   BackColor       =   &H0080FF80&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Database Management"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4905
   BeginProperty Font 
      Name            =   "Segoe Print"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmdata_man.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   4905
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Restore"
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   6360
      Width           =   4695
      Begin VB.ComboBox Combo1 
         Height          =   675
         Left            =   2160
         TabIndex        =   11
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Restore Database"
         Height          =   1095
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1680
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select table"
         Height          =   555
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Reset"
      Height          =   3735
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   4695
      Begin VB.TextBox Text2 
         Height          =   675
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1560
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   675
         Left            =   2040
         TabIndex        =   5
         Top             =   600
         Width           =   2535
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Reset"
         Height          =   915
         Left            =   1440
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   555
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "User name"
         Height          =   555
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Backup"
      Height          =   2055
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Back up Database"
         Height          =   1215
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   600
         Width           =   2775
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   1920
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
End
Attribute VB_Name = "frmdata_man"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Dim table(10) As String
Dim tab_col(10) As Integer
Dim t2 As New DataSource
Private Sub Command1_Click()
On Error GoTo cdhandler
start:
With CD
    .DialogTitle = "Back Up Database"
    .DefaultExt = "mdb"
    .Filter = "MS Access Document(*.mdb)|*.mdb"
    .Flags = cdlOFNOverwritePrompt
    .FileName = App.Path & "\back up\all_" & Format(Date, "dd-mmm-yy") & ".mdb"
    .ShowSave
End With
Dim fpath As New FileSystemObject
fpath.GetFile(App.Path & "\all.mdb").Copy CD.FileName, True
MsgBox "Database Back Up Successful", vbInformation
Exit Sub
cdhandler:
If Not Err.Number = 32755 Then
    Select Case MsgBox(Error(Err.Number), vbCritical + vbAbortRetryIgnore, "Error number-" & Str(Err.Number))
        Case vbAbort
            Resume exitline
        Case vbRetry
            Resume start
        Case vbIgnore
            Resume Next
        End Select
End If
exitline:
End Sub

Private Sub Command2_Click()
If Text1.Text = "" Or Text2.Text = "" Then
    MsgBox "Enter Username And Password", vbExclamation
    Exit Sub
End If
m = MsgBox("Are You Sure ?", vbQuestion + vbYesNo)
If m = vbYes Then
    t.Database ("Select User,Password from Admin Where Type='Administrator'")
    If t.rs.EOF Then
        MsgBox "please check Username And Password", vbExclamation
        Text1.Text = ""
        Text2.Text = ""
        Text1.SetFocus
    Else
        t.rs.Close
            t.db.Execute ("Delete from Admin where Type not like 'Admin%'")
        For i = 1 To 4
            t.db.Execute ("Delete from " & table(i))
        Next
            t.db.Execute ("Delete from " & table(5) & "where ID not like '1'")
            t.db.Execute ("Delete from " & table(6))
        Dim fpath As New FileSystemObject
            If fpath.FolderExists(App.Path & "\vo_img") Then: Call fpath.DeleteFolder(App.Path & "\vo_img", True)
            fpath.CreateFolder (App.Path & "\vo_img")
            If fpath.FolderExists(App.Path & "\party_logo") Then Call fpath.DeleteFolder(App.Path & "\party_logo", True)
            fpath.CreateFolder (App.Path & "\party_logo")
    End If
End If
End Sub

Private Sub Command3_Click()
On Error GoTo exit1
start:
If Combo1.Text = "" Then
    MsgBox "please select a table", vbExclamation
    Exit Sub
End If
With CD
    .DialogTitle = "Restore Database"
    .Filter = "MS Access Document(*.mdb)|*.mdb"
    .ShowOpen
    .Flags = cdlOFNHideReadOnly
End With
    If Combo1.Text <> "All" Then
        m = t2.Database(CD.FileName, "Select * from " & table(Combo1.ListIndex))
        t.Database ("Select * from " & table(Combo1.ListIndex))
        t.db.Execute ("Delete from " & table(Combo1.ListIndex))
        t2.rs.MoveFirst
        While Not t2.rs.EOF
                t.rs.AddNew
                For i = 0 To tab_col(Combo1.ListIndex) - 1
                    t.rs.Fields(i) = t2.rs.Fields(i)
                Next
                t.rs.Update
                t2.rs.MoveNext
        Wend
        t.rs.Close
        t.db.Close
        t2.rs.Close
        t2.db.Close
        MsgBox "Data has been sucessfully restored", vbInformation
    Else
        Dim fpath As New FileSystemObject
        fpath.GetFile(App.Path & "\all.mdb").Copy CD.FileName, True
        MsgBox "Database has been sucessfully restored", vbInformation
    End If
Exit Sub
exit1:
If Not Err.Number = 32755 Then
    Select Case MsgBox(Error(Err.Number), vbCritical + vbAbortRetryIgnore, "Error number-" & Str(Err.Number))
        Case vbAbort
            Resume exitline
        Case vbRetry
            Resume start
        Case vbIgnore
            Resume Next
        End Select
End If
exitline:
End Sub

Private Sub Form_Load()
table(0) = "Admin": tab_col(0) = 5
table(1) = "Voter": tab_col(1) = 8
table(2) = "Candidate": tab_col(2) = 7
table(5) = "Party": tab_col(5) = 3
table(3) = "Post": tab_col(3) = 2
table(4) = "Courses": tab_col(4) = 3
table(6) = "Winner": tab_col(6) = 7
table(7) = "backup": tab_col(7) = 1
For i = 0 To 7
    Combo1.AddItem table(i)
Next
Combo1.AddItem "All"
End Sub
