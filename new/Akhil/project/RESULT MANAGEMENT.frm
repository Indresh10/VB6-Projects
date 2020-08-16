VERSION 5.00
Begin VB.Form splash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8205
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10950
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Viner Hand ITC"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "RESULT MANAGEMENT.frx":0000
   ScaleHeight     =   8205
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   4000
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " MANAGEMENT"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   8160
      TabIndex        =   3
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "GUIDE BY:- ASHIM SIR"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   7200
      Width           =   4095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "AKHIL AJMERA"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   5280
      TabIndex        =   1
      Top             =   7200
      Width           =   4095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   " DIGITAL WAY OF"
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   -120
      TabIndex        =   0
      Top             =   1320
      Width           =   3495
   End
End
Attribute VB_Name = "splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()
login.Visible = True
Unload Me
End Sub
