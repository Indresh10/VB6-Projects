VERSION 5.00
Begin VB.Form frmsplash 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8970
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11790
   BeginProperty Font 
      Name            =   "MV Boli"
      Size            =   26.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   8970
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   5280
      Top             =   4200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8950
      Left            =   0
      Picture         =   "frmsplash.frx":0000
      ScaleHeight     =   8925
      ScaleWidth      =   11745
      TabIndex        =   0
      Top             =   0
      Width           =   11775
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Press Any Key or click here to continue"
         ForeColor       =   &H000000FF&
         Height          =   690
         Left            =   690
         TabIndex        =   3
         Top             =   8160
         Visible         =   0   'False
         Width           =   10515
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Loading"
         ForeColor       =   &H000000FF&
         Height          =   690
         Left            =   4845
         TabIndex        =   2
         Top             =   8160
         Width           =   1920
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H008080FF&
         Caption         =   "Label1"
         ForeColor       =   &H00FFFFFF&
         Height          =   690
         Left            =   720
         TabIndex        =   1
         Top             =   1150
         Width           =   1545
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   2520
   End
   Begin VB.TextBox Text1 
      Height          =   810
      Left            =   5280
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4200
      Width           =   1215
   End
End
Attribute VB_Name = "frmsplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim count1 As Single
Dim per As Integer
Dim col1, col2 As Long
Private Sub Form_Load()
pi = 3.14159
start = (pi / 2)
count1 = start + 0.0628319
per = 0
col1 = 0
col2 = 255
End Sub

Private Sub Label3_Click()
Login.Show
Unload Me
End Sub

Private Sub Picture1_Click()
If Label3.Visible = True Then Text1.SetFocus
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
Login.Show: Unload Me
End Sub

Private Sub Timer1_Timer()
pi = 3.14159
deg = 0.0628319
Picture1.FillStyle = 0
Select Case per
Case 0 To 50
    Picture1.FillColor = RGB(col2, col1, 0)
Case 51 To 100
    Picture1.FillColor = RGB(0, col2, col1)
End Select
col1 = col1 + 5
col2 = col2 - 5
If col2 = 0 Then col1 = 0: col2 = 255
Picture1.DrawWidth = 10
count1 = count1 + deg
If count1 = 6.346018 Then count1 = 0
per = per + 1
Label1.Caption = per & "%"
Label2.Caption = Label2.Caption + "."
If per Mod 3 = 0 Then Label2.Caption = "Loading"
Picture1.Circle (1500, 1500), 1200, vbRed, -(pi / 2), -(count1)
If per = 100 Then
    Timer1.Enabled = False
    Picture1.Cls
    Picture1.Circle (1500, 1500), 1200, vbRed
    Label2.Visible = False
    Label3.Visible = True
    Text1.SetFocus
    Timer2.Enabled = True
    per = 1
    col2 = 255: col1 = 0
    Label1.BackStyle = vbTransparent
End If
End Sub

Private Sub Timer2_Timer()
    If per = 1 Then
    Picture1.FillColor = RGB(255, col1, 0)
    Label3.ForeColor = RGB(255, col1, 0)
        If col1 = 255 Then per = 2: col1 = 0: col2 = 255
    ElseIf per = 2 Then
    Picture1.FillColor = RGB(col2, 255, 0)
    Label3.ForeColor = RGB(col2, 255, 0)
        If col2 = 0 Then per = 3: col1 = 0
    ElseIf per = 3 Then
    Picture1.FillColor = RGB(0, 255, col1)
    Label3.ForeColor = RGB(0, 255, col1)
        If col1 = 255 Then per = 4: col1 = 0: col2 = 255
    ElseIf per = 4 Then
    Picture1.FillColor = RGB(0, col2, 255)
    Label3.ForeColor = RGB(0, col2, 255)
        If col2 = 0 Then per = 5: col1 = 0
    ElseIf per = 5 Then
    Picture1.FillColor = RGB(col1, 0, 255)
    Label3.ForeColor = RGB(col1, 0, 255)
        If col1 = 255 Then per = 6: col1 = 0: col2 = 255
    ElseIf per = 6 Then
    Picture1.FillColor = RGB(255, 0, col2)
    Label3.ForeColor = RGB(255, 0, col2)
        If col2 = 0 Then per = 1: col1 = 0
    End If
Picture1.Circle (1500, 1500), 1200, vbRed
col1 = col1 + 5
col2 = col2 - 5
End Sub
