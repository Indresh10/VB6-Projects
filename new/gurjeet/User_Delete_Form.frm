VERSION 5.00
Begin VB.Form FrmUserDelete 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Deletion"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   Icon            =   "User_Delete_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete"
      Default         =   -1  'True
      Height          =   375
      Left            =   2580
      TabIndex        =   3
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdBack 
      Caption         =   "&Back"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4800
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.ListBox LstUserDelete 
      Height          =   2595
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label LblUserSelect 
      AutoSize        =   -1  'True
      Caption         =   "&Select User to delete :"
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
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2145
   End
End
Attribute VB_Name = "FrmUserDelete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_user As New ADODB.Recordset
Dim rs_tmp As New ADODB.Recordset

Private Sub CmdBack_Click()
    Unload Me
    FrmUserMng.Show vbModal
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdDelete_Click()
    Dim Query As String, cnt As Integer
        
    'CHECK FOR RECORDSET IS OPEN OR CLOSED
    If rs_tmp.State = 1 Then
        rs_tmp.Close
    End If
    
    rs_tmp.Open "select * from Login_Mast where Typ='A'", conn, adOpenStatic, adLockPessimistic
    
    If (rs_tmp.RecordCount = 1) And (Mid(LstUserDelete.Text, Len(LstUserDelete.Text) - 1, 1) = "A") Then
        MsgBox "You can not delete this Admin user." & vbCrLf & "Atlist one Admin user is required.", vbCritical, "User Deletion"
        Exit Sub
    End If
    
    If MsgBox("You want to delete selected user ?", vbQuestion + vbOKCancel, "User Deletion") = vbOK Then
        'FIND SELECTED USER
        rs_user.MoveFirst
        rs_user.Find "usr='" & Mid(LstUserDelete.Text, 1, Len(LstUserDelete.Text) - 4) & "'"
    
        rs_user.Delete  'DELETE USER
    
        Call fillList   'FILL LIST BOX
    End If
End Sub

Private Sub Form_Load()
    MDIFrm.Pct1.Visible = False
    'OPEN RECORDSET
    rs_user.Open "select * from Login_Mast", conn, adOpenStatic, adLockPessimistic
    
    Call fillList   'FILL LIST BOX
End Sub

Private Sub Form_Unload(Cancel As Integer)
    rs_user.Close
End Sub

Private Sub fillList()
    'FILL ListBox
    LstUserDelete.Clear
    If rs_user.RecordCount <> 0 Then
        rs_user.MoveFirst
        While Not rs_user.EOF
            LstUserDelete.AddItem rs_user.Fields(0) & " (" & rs_user.Fields(2) & ")"
            rs_user.MoveNext
        Wend
        LstUserDelete.Text = LstUserDelete.List(0)
    End If
End Sub
