VERSION 5.00
Begin VB.Form FrmWelcome 
   Caption         =   "Welcome Form"
   ClientHeight    =   10935
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   20250
   Icon            =   "Library.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   165
      Top             =   255
   End
   Begin VB.Label LblSys2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   1155
      Left            =   7943
      TabIndex        =   7
      Top             =   7890
      Width           =   3660
   End
   Begin VB.Label LblSys1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1155
      Left            =   7883
      TabIndex        =   6
      Top             =   7935
      Width           =   3660
   End
   Begin VB.Label LblLib2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00226622&
      Height          =   1155
      Left            =   5138
      TabIndex        =   5
      Top             =   6075
      Width           =   10035
   End
   Begin VB.Label LblLib1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Library Management"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   1155
      Left            =   5078
      TabIndex        =   4
      Top             =   6105
      Width           =   10035
   End
   Begin VB.Label LblRpbc2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1155
      Left            =   9533
      TabIndex        =   3
      Top             =   4275
      Width           =   1185
   End
   Begin VB.Label LblRpbc1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "To"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   1155
      Left            =   9533
      TabIndex        =   2
      Top             =   4230
      Width           =   1185
   End
   Begin VB.Label LblWelcome1 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1605
      Left            =   6488
      TabIndex        =   1
      Top             =   1905
      Width           =   6945
   End
   Begin VB.Label LblWelcome2 
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   6563
      TabIndex        =   0
      Top             =   1845
      Width           =   6900
   End
   Begin VB.Image Img1 
      Height          =   11415
      Left            =   0
      Picture         =   "Library.frx":0ECA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   20535
   End
End
Attribute VB_Name = "FrmWelcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
    FrmLogin.Show vbModal
    Timer1.Enabled = False
End Sub

