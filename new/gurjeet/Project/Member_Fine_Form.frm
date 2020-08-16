VERSION 5.00
Begin VB.Form FrmMbrFine 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Fine"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox LstMember 
      Height          =   2010
      Left            =   1680
      TabIndex        =   15
      Top             =   2520
      Width           =   4935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1080
      TabIndex        =   13
      Top             =   5160
      Width           =   975
   End
   Begin VB.TextBox TxtCode 
      Height          =   375
      Left            =   1680
      MaxLength       =   6
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.TextBox TxtSurname 
      Height          =   375
      Left            =   1680
      TabIndex        =   4
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox TxtFee 
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   1913
      Width           =   1215
   End
   Begin VB.TextBox TxtFirst 
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.TextBox TxtLast 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label LblMember 
      AutoSize        =   -1  'True
      Caption         =   "Select &Member :"
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
      Left            =   240
      TabIndex        =   12
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label LblCode 
      AutoSize        =   -1  'True
      Caption         =   "Co&de :"
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
      Left            =   240
      TabIndex        =   11
      Top             =   780
      Width           =   585
   End
   Begin VB.Label LblName 
      AutoSize        =   -1  'True
      Caption         =   "Na&me :"
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
      Left            =   240
      TabIndex        =   10
      Top             =   1500
      Width           =   645
   End
   Begin VB.Label LblFee 
      AutoSize        =   -1  'True
      Caption         =   "Fine &Rupees :"
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
      Left            =   240
      TabIndex        =   9
      Top             =   1980
      Width           =   1245
   End
   Begin VB.Label LblSurname 
      AutoSize        =   -1  'True
      Caption         =   "Surname"
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
      Left            =   1920
      TabIndex        =   8
      Top             =   1200
      Width           =   810
   End
   Begin VB.Label LblFirst 
      AutoSize        =   -1  'True
      Caption         =   "Member Name"
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
      Left            =   3120
      TabIndex        =   7
      Top             =   1200
      Width           =   1350
   End
   Begin VB.Label LblLast 
      AutoSize        =   -1  'True
      Caption         =   "Father Name"
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
      Left            =   5040
      TabIndex        =   6
      Top             =   1200
      Width           =   1170
   End
   Begin VB.Label LblLabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "MEMBER FINE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   1710
      TabIndex        =   0
      Top             =   75
      Width           =   3450
   End
   Begin VB.Shape ShapLabel 
      FillColor       =   &H008080FF&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   0
      Width           =   6840
   End
End
Attribute VB_Name = "FrmMbrFine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
