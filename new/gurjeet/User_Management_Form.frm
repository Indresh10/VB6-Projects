VERSION 5.00
Begin VB.Form FrmUserMng 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Account"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   FillStyle       =   0  'Solid
   Icon            =   "User_Management_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   4950
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdDeleteAcc 
      Caption         =   "&Delete account"
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
      Left            =   960
      TabIndex        =   2
      Top             =   2160
      Width           =   2895
   End
   Begin VB.CommandButton CmdEditAcc 
      Caption         =   "&Edit your account"
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
      Left            =   960
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
   Begin VB.CommandButton CmdCreateAcc 
      Caption         =   "Create a &new account"
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
      Left            =   960
      TabIndex        =   0
      Top             =   480
      Width           =   2895
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3720
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "FrmUserMng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCreateAcc_Click()
    Unload Me
    FrmCreateAcc.Show vbModal
End Sub

Private Sub CmdDeleteAcc_Click()
    Unload Me
    FrmUserDelete.Show vbModal
End Sub

Private Sub CmdEditAcc_Click()
    Unload Me
    FrmEditAcc.Show vbModal
End Sub

Private Sub Form_Load()
    If userType = "L" Then
        CmdCreateAcc.Enabled = False
        CmdDeleteAcc.Enabled = False
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Forms.Count = 2 Then
        MDIFrm.Pct1.Visible = True
    End If
End Sub
