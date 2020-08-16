VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form Animation 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9030
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Play"
      Height          =   495
      Left            =   3600
      TabIndex        =   1
      Top             =   3240
      Width           =   1695
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   4683
      _Version        =   327681
      FullWidth       =   545
      FullHeight      =   177
   End
End
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Animation1.Play
End Sub

Private Sub Form_Load()
Animation1.Open ("C:\Windows\WinSxS\amd64_microsoft-windows-tabletpc-inputpanel_31bf3856ad364e35_10.0.15063.0_none_22a2eeffb0510686\join.avi")
End Sub

