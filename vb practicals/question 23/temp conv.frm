VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4050
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Conversion"
      Height          =   1935
      Left            =   1800
      TabIndex        =   2
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton Option2 
         Caption         =   "Kelvin"
         Height          =   495
         Left            =   360
         TabIndex        =   4
         Top             =   1200
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Farheniet"
         Height          =   495
         Left            =   360
         TabIndex        =   3
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4560
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   " "
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   3360
      Width           =   105
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Converted temp"
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   3360
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Enter Degree in centigrade"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   3870
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Option1_Click()
c = Val(Text1.Text)
If Option1.Value = True Then f = (c * 9 / 5) + 32
Label3.Caption = f
Label2.Caption = "Converted temp in farheniet"
End Sub

Private Sub Option2_Click()
c = Val(Text1.Text)
If Option2.Value = True Then k = c + 273.15
Label3.Caption = k
Label2.Caption = "Converted temp in kelvin"
End Sub
