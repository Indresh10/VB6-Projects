VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6420
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9240
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   18
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6420
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check4 
      Caption         =   "Strike through"
      Height          =   615
      Left            =   1320
      TabIndex        =   11
      Top             =   4200
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   615
      Left            =   4440
      TabIndex        =   9
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset"
      Height          =   615
      Left            =   6480
      TabIndex        =   8
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sample"
      Height          =   2415
      Left            =   5640
      TabIndex        =   7
      Top             =   1920
      Width           =   2655
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aa Xx Yy"
         Height          =   435
         Left            =   585
         TabIndex        =   10
         Top             =   1080
         Width           =   1455
      End
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Underline"
      Height          =   435
      Left            =   1320
      TabIndex        =   4
      Top             =   3440
      Width           =   2535
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Italic"
      Height          =   435
      Left            =   1320
      TabIndex        =   3
      Top             =   2805
      Width           =   1215
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bold"
      Height          =   435
      Left            =   1320
      TabIndex        =   2
      Top             =   2160
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   555
      Left            =   5640
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.ComboBox Combo1 
      Height          =   555
      Left            =   1320
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Select font size"
      Height          =   435
      Left            =   5640
      TabIndex        =   6
      Top             =   480
      Width           =   2430
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Select font"
      Height          =   435
      Left            =   1320
      TabIndex        =   5
      Top             =   480
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
If Check1.Value = 1 Then
    Label3.FontBold = True
Else
    Label3.FontBold = False
End If

End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Label3.FontItalic = True
Else
    Label3.FontItalic = False
End If
End Sub

Private Sub Check3_Click()
If Check3.Value = 1 Then
    Label3.FontUnderline = True
Else
    Label3.FontUnderline = False
End If
End Sub
Private Sub Check4_Click()
If Check4.Value = 1 Then
    Label3.FontStrikethru = True
Else
    Label3.FontStrikethru = False
End If
End Sub

Private Sub Combo1_Click()
Label3.Font = Combo1.Text
End Sub

Private Sub Combo2_Change()
Label3.FontSize = Combo2.Text
End Sub

Private Sub Combo2_Click()
Label3.FontSize = Combo2.Text
End Sub

Private Sub Combo2_Validate(Cancel As Boolean)
If Val(Combo2.Text) > 72 Then
MsgBox "Size not avalable"
Combo2.Text = "8"
End If
End Sub

Private Sub Command1_Click()
Combo1.ListIndex = 1
Combo2.ListIndex = 1
Check1.Value = 0
Check2.Value = 0
Check3.Value = 0
Check4.Value = 0
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Form_Load()
For i = 1 To Screen.FontCount
    Combo1.AddItem Screen.Fonts(i)
Next
For i = 8 To 24 Step 2
    Combo2.AddItem i
Next
Combo1.ListIndex = 1
Combo2.ListIndex = 0
End Sub
