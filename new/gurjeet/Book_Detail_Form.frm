VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBkDtl 
   Caption         =   "Book Detail"
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3300
   ScaleWidth      =   7320
   WindowState     =   2  'Maximized
   Begin MSFlexGridLib.MSFlexGrid MSFlxGrd 
      Height          =   2535
      Left            =   180
      TabIndex        =   3
      Top             =   150
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   4471
      _Version        =   393216
   End
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdDelete 
      Caption         =   "&Delete Book"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton CmdEnter 
      Caption         =   "&Enter Book"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "FrmBkDtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    If Me.Width > 7440 And Me.Height > 3705 Then
        'POSITION THE FLEXGRID
        MSFlxGrd.Width = Me.ScaleWidth - 200
        MSFlxGrd.Height = Me.ScaleHeight - 1000
    
        'POSITION COMMAND BUTTONS
        CmdEnter.Top = MSFlxGrd.Height + 400
        CmdDelete.Top = CmdEnter.Top
        CmdCancel.Top = CmdEnter.Top
        CmdCancel.Left = Me.ScaleWidth - CmdCancel.Width - 500
    End If
End Sub
