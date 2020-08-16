VERSION 5.00
Begin VB.MDIForm MDIForm1 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Election Management System"
   ClientHeight    =   10935
   ClientLeft      =   -45
   ClientTop       =   465
   ClientWidth     =   20250
   Icon            =   "base.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3600
      Top             =   4560
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   4680
   End
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   3360
      Top             =   4560
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11520
      Left            =   0
      LinkTimeout     =   0
      Picture         =   "base.frx":1542
      ScaleHeight     =   11520
      ScaleWidth      =   20250
      TabIndex        =   0
      Top             =   0
      Width           =   20250
      Begin VB.PictureBox Picture2 
         BackColor       =   &H8000000D&
         Height          =   11520
         Left            =   0
         ScaleHeight     =   11460
         ScaleWidth      =   3660
         TabIndex        =   2
         Top             =   0
         Width           =   3720
         Begin VB.Frame Frame1 
            BackColor       =   &H8000000D&
            Caption         =   "Welcome"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2655
            Left            =   120
            TabIndex        =   8
            Top             =   4440
            Width           =   3375
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "MV Boli"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   570
               Left            =   915
               TabIndex        =   9
               Top             =   1200
               Width           =   1395
            End
            Begin VB.Label Label9 
               Caption         =   "Label9"
               Height          =   375
               Left            =   1440
               TabIndex        =   14
               Top             =   1320
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label Label10 
               Caption         =   "Label10"
               Height          =   135
               Left            =   1800
               TabIndex        =   15
               Top             =   1560
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H8000000D&
            Caption         =   "Date and Time"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   3375
            Left            =   120
            TabIndex        =   3
            Top             =   7440
            Width           =   3375
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Date"
               BeginProperty Font 
                  Name            =   "MV Boli"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   510
               Left            =   240
               TabIndex        =   7
               Top             =   600
               Width           =   975
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "MV Boli"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   240
               TabIndex        =   6
               Top             =   1200
               Width           =   1305
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               Caption         =   "Label3"
               BeginProperty Font 
                  Name            =   "MV Boli"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   510
               Left            =   240
               TabIndex        =   5
               Top             =   2520
               Width           =   1305
            End
            Begin VB.Label Label7 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackColor       =   &H00FFFFFF&
               BackStyle       =   0  'Transparent
               Caption         =   "Time"
               BeginProperty Font 
                  Name            =   "MV Boli"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   510
               Left            =   240
               TabIndex        =   4
               Top             =   1800
               Width           =   975
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "UserAccount Panel "
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   465
            Left            =   120
            TabIndex        =   13
            Top             =   0
            Width           =   3585
         End
         Begin VB.Image Image1 
            Height          =   960
            Left            =   360
            Picture         =   "base.frx":1A188
            Top             =   720
            Width           =   960
         End
         Begin VB.Image Image2 
            Height          =   975
            Left            =   360
            Picture         =   "base.frx":1ABC0
            Stretch         =   -1  'True
            Top             =   3360
            Width           =   975
         End
         Begin VB.Image Image3 
            Height          =   960
            Left            =   360
            Picture         =   "base.frx":1B281
            Top             =   2100
            Width           =   960
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " Change Password"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1095
            Index           =   0
            Left            =   1320
            TabIndex        =   12
            Top             =   720
            Width           =   2100
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Logout"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   510
            Index           =   1
            Left            =   1800
            TabIndex        =   11
            Top             =   2280
            Width           =   1380
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Exit"
            BeginProperty Font 
               Name            =   "MV Boli"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   510
            Index           =   2
            Left            =   1965
            TabIndex        =   10
            Top             =   3600
            Width           =   810
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00FFFFFF&
            BorderWidth     =   5
            X1              =   0
            X2              =   3700
            Y1              =   480
            Y2              =   480
         End
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "About The Developer"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   10920
         TabIndex        =   22
         Top             =   5880
         Width           =   4035
      End
      Begin VB.Image Image10 
         Height          =   2250
         Left            =   11760
         Picture         =   "base.frx":1E2AE
         Top             =   7200
         Width           =   2250
      End
      Begin VB.Image Image9 
         Height          =   3360
         Left            =   16800
         Picture         =   "base.frx":1F13C
         Stretch         =   -1  'True
         Top             =   2280
         Width           =   3360
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Import data from College Database"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   16320
         TabIndex        =   21
         Top             =   840
         Width           =   4035
      End
      Begin VB.Image Image8 
         Height          =   3000
         Left            =   7800
         Picture         =   "base.frx":20AAA
         Stretch         =   -1  'True
         Top             =   6960
         Width           =   2880
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Candidates,Parties Maintainance"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   10815
         TabIndex        =   19
         Top             =   720
         Width           =   4035
      End
      Begin VB.Image Image4 
         Height          =   3360
         Left            =   11040
         Picture         =   "base.frx":25504
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   3360
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Database Management"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   14400
         TabIndex        =   18
         Top             =   5520
         Visible         =   0   'False
         Width           =   4035
      End
      Begin VB.Image Image7 
         Height          =   3840
         Left            =   14520
         Picture         =   "base.frx":29294
         Stretch         =   -1  'True
         Top             =   6720
         Visible         =   0   'False
         Width           =   3720
      End
      Begin VB.Image Image6 
         Height          =   960
         Left            =   18000
         Picture         =   "base.frx":2B735
         Top             =   10200
         Width           =   960
      End
      Begin VB.Label Label 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Christ College Jagdalpur"
         BeginProperty Font 
            Name            =   "Segoe Print"
            Size            =   27.75
            Charset         =   161
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   975
         Left            =   3840
         TabIndex        =   1
         Top             =   -120
         Width           =   6975
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFC0&
         BackStyle       =   0  'Transparent
         Caption         =   "Courses"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   510
         Left            =   18960
         TabIndex        =   17
         Top             =   10440
         Width           =   1515
      End
      Begin VB.Image Image5 
         Height          =   3120
         Left            =   4320
         Picture         =   "base.frx":2BBE9
         Stretch         =   -1  'True
         Top             =   2160
         Width           =   3120
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   10
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   5775
         Left            =   3960
         Shape           =   3  'Circle
         Top             =   840
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Voter's Maintainence"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   3840
         TabIndex        =   16
         Top             =   720
         Width           =   4035
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   10
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   5775
         Left            =   10800
         Shape           =   3  'Circle
         Top             =   960
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Shape Shape3 
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   975
         Left            =   17880
         Top             =   10200
         Width           =   2775
      End
      Begin VB.Shape Shape4 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   10
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   5175
         Left            =   14520
         Shape           =   3  'Circle
         Top             =   6000
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   10
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   5535
         Left            =   7320
         Shape           =   3  'Circle
         Top             =   5640
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "Votes Standing"
         BeginProperty Font 
            Name            =   "MV Boli"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1095
         Left            =   7200
         TabIndex        =   20
         Top             =   5760
         Width           =   4035
      End
      Begin VB.Shape Shape6 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   10
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   5775
         Left            =   16545
         Shape           =   3  'Circle
         Top             =   1080
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.Shape Shape7 
         BorderColor     =   &H0000FFFF&
         BorderWidth     =   10
         FillColor       =   &H00FF8080&
         FillStyle       =   0  'Solid
         Height          =   5895
         Left            =   11400
         Shape           =   3  'Circle
         Top             =   5400
         Visible         =   0   'False
         Width           =   2895
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ex As Integer

Private Sub Image10_DblClick()
frmabout.Show vbModal, Me
End Sub

Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape7.Visible = True
End Sub

Private Sub Image4_DblClick()
frmcand_man.Show vbModal, MDIForm1
End Sub

Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = True
End Sub

Private Sub Image5_DblClick()
frmvoter.Show vbModal, MDIForm1
End Sub

Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape2.Visible = True
End Sub

Private Sub Image7_DblClick()
frmdata_man.Show vbModal, MDIForm1
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape4.Visible = True
End Sub

Private Sub Image8_DblClick()
frmvote_stand.Show vbModal, MDIForm1
End Sub

Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape5.Visible = True
End Sub
Private Sub Image9_DblClick()
frmclg_data.Show vbModal, MDIForm1
End Sub

Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape6.Visible = True
End Sub

Private Sub Label12_Click()
frmcourses.Show vbModal, MDIForm1
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label12.ForeColor = vbRed
End Sub

Private Sub Label2_Click(Index As Integer)
Select Case Index
Case 0
    frmchng_pass.Show
    frmchng_pass.Text1.Visible = True
    frmchng_pass.Label1.Visible = True
Case 1
    ex = 0
    Load Login
    Unload Me
    Login.Show
    MsgBox "Logout Successfull", vbInformation, "Election Management system"
    Login.Combo1.Text = ""
    Login.Text1.Text = ""
    Login.Text2.Text = ""
Case 2
    End
End Select
End Sub

Private Sub Label2_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
    Label2(0).ForeColor = vbYellow
    Image1.Picture = LoadPicture(App.Path & "\logo\pass_reset.jpg")
Case 1
    Label2(1).ForeColor = vbYellow
    Image3.Picture = LoadPicture(App.Path & "\logo\logout.jpg")
Case 2
    Label2(2).ForeColor = vbYellow
    Image2.Picture = LoadPicture(App.Path & "\logo\exit2.gif")
End Select
End Sub

Private Sub MDIForm_Load()
ex = 1
Label5.Caption = Format(Date, "dd-mmm-yy")
Label6.Caption = Format(Time, "hh:mm:ss am/pm")
Label3.Caption = Login.Text1.Text
MDIForm1.Icon = LoadPicture(App.Path & "\logo\election.ico")
If Login.Combo1.Text = "Administrator" Then: Image7.Visible = True: Label13.Visible = True
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
If Not ex = 0 Then End
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Shape1.Visible = False
Shape2.Visible = False
Shape4.Visible = False
Shape5.Visible = False
Shape6.Visible = False
Shape7.Visible = False
Label12.ForeColor = vbBlack
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2(0).ForeColor = vbWhite
Image1.Picture = LoadPicture(App.Path & "\logo\passreset.jpg")
Label2(1).ForeColor = vbWhite
Image2.Picture = LoadPicture(App.Path & "\logo\exit.gif")
Label2(2).ForeColor = vbWhite
Image3.Picture = LoadPicture(App.Path & "\logo\logout1.jpg")
End Sub

Private Sub Timer1_Timer()
Label6.Caption = Format(Time, "hh:mm:ss am/pm")
End Sub

Private Sub Timer2_Timer()
Label.ForeColor = vbCyan
Label.Left = Label.Left + 10
If Label.Left = 13470 Then
    Timer2.Enabled = False
    Timer3.Enabled = True
End If
End Sub

Private Sub Timer3_Timer()
Label.ForeColor = vbYellow
Label.Left = Label.Left - 10
If Label.Left = 3840 Then
    Timer2.Enabled = True
    Timer3.Enabled = False
End If
End Sub
