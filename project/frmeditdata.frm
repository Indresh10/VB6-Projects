VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmeditdata 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Edit Data"
   ClientHeight    =   8310
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12030
   BeginProperty Font 
      Name            =   "Lucida Handwriting"
      Size            =   15.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmeditdata.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   12030
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   8760
      ScaleHeight     =   3345
      ScaleWidth      =   3105
      TabIndex        =   10
      Top             =   4800
      Width           =   3135
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Generate Voter ID card"
         Height          =   1095
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Update Data And Exit"
         Height          =   1095
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Options"
         Height          =   495
         Left            =   240
         TabIndex        =   13
         Top             =   120
         Width           =   2655
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   8760
      ScaleHeight     =   4545
      ScaleWidth      =   3105
      TabIndex        =   8
      Top             =   120
      Width           =   3135
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Upload Photo"
         Height          =   1095
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3240
         Width           =   2895
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Photo"
         Height          =   495
         Left            =   240
         TabIndex        =   14
         Top             =   120
         Width           =   2655
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2295
         Left            =   600
         Stretch         =   -1  'True
         Top             =   840
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   8055
      Left            =   120
      ScaleHeight     =   8025
      ScaleWidth      =   8505
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   5160
         TabIndex        =   22
         Top             =   4920
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Format          =   123076609
         CurrentDate     =   43817
      End
      Begin VB.ComboBox Combo2 
         Height          =   525
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   3960
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         Height          =   525
         Left            =   5160
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   3000
         Width           =   2415
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         Height          =   2175
         Left            =   2880
         TabIndex        =   16
         Top             =   5640
         Width           =   3255
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Change"
            Height          =   615
            Left            =   645
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1320
            Width           =   2055
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Show"
            Height          =   615
            Left            =   645
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   480
            Width           =   2055
         End
      End
      Begin VB.TextBox Text2 
         Height          =   615
         Left            =   5160
         TabIndex        =   1
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         Height          =   525
         Left            =   4320
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   6480
         Width           =   375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Edit And Confirm Data"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   8295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Voter ID"
         Height          =   405
         Left            =   1080
         TabIndex        =   7
         Top             =   1320
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   405
         Left            =   1080
         TabIndex        =   6
         Top             =   2130
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class"
         Height          =   405
         Left            =   1080
         TabIndex        =   5
         Top             =   3060
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Year"
         Height          =   405
         Left            =   1080
         TabIndex        =   4
         Top             =   3990
         Width           =   900
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date Of Birth "
         Height          =   405
         Left            =   1080
         TabIndex        =   3
         Top             =   5040
         Width           =   2460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "[Voter ID] "
         Height          =   405
         Left            =   5160
         TabIndex        =   2
         Top             =   1320
         Width           =   1875
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   5760
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmeditdata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim t As New Database1
Private Sub Command1_Click()
On Error GoTo cdhandler
start:
t.Database ("select * from Voter where V_ID='" + Label6.Caption + "'")
With CD
    .Filter = "JPEG Files(*.jpg)|*.jpg|GIF Files(*.gif)|*.gif"
    .ShowOpen
End With
fn = "test.jpg"
If Not CD.FileName = "" Then
FileCopy CD.FileName, App.Path & "\vo_img\" + fn
file = App.Path & "\vo_img\" + fn
t.rs.Update "file", fn
Image1.Picture = LoadPicture(file)
End If
Exit Sub
cdhandler:
    Select Case MsgBox(Error(Err.Number), vbCritical + vbAbortRetryIgnore, "Error number-" & Str(Err.Number))
    Case vbAbort
        Resume exitline
    Case vbRetry
        Resume start
    Case vbIgnore
        Resume Next
    End Select
exitline:
End Sub

Private Sub Command2_Click()
If Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then MsgBox "Please Check your details", vbExclamation: Exit Sub
Call save
Unload Me
End Sub

Private Sub Command3_Click()
If Text2.Text = "" Or Combo1.Text = "" Or Combo2.Text = "" Then MsgBox "Please Check your details", vbExclamation: Exit Sub
Call save
t.Database ("select * from Voter where V_ID='" + Label6.Caption + "'")
Set IDreport.DataSource = t.rs
IDreport.Show vbModal, Me
End Sub

Private Sub Command4_Click()
MsgBox "Your password is " + Text1.Text, vbInformation
End Sub

Private Sub Command5_Click()
frmchng_pass.Show vbModal, Me
End Sub

Private Sub Form_Load()
Label6.Caption = Login.Text5.Text
t.Database ("Select * from Voter where V_ID='" + Label6.Caption + "'")
With t.rs
    Text2.Text = .Fields(2)
    Combo1.Text = .Fields(3)
    Combo2.Text = .Fields(4)
    dt = Format(.Fields(5), "mm/dd/yyyy")
    DTPicker1.Value = dt
    Text1.Text = .Fields(6)
End With
Dim rs As New ADODB.Recordset
rs.Open "Select Name from Courses", t.db, adOpenKeyset, adLockOptimistic
rs.MoveFirst
While Not rs.EOF
    Combo1.AddItem rs.Fields(0)
    rs.MoveNext
Wend
rs.Close
Combo2.AddItem "First"
Combo2.AddItem "Second"
Combo2.AddItem "Third"
If IsNull(t.rs.Fields(7)) Then t.rs.Update "file", "base.gif"
Image1.Picture = LoadPicture(App.Path & "\vo_img\" & t.rs.Fields(7))
End Sub

Private Function save()
t.Database ("select * from Voter where V_ID='" + Label6.Caption + "'")
If t.rs.Fields("file") = "base.gif" Then
file = "base.gif"
Else
file = Left$(Text2.Text, 3) & t.rs.Fields(0) & ".jpg"
Name App.Path & "\vo_img\" & t.rs.Fields("file") As App.Path & "\vo_img\" & file
End If
With t.rs
    .Fields(2) = Text2.Text
    .Fields(3) = Combo1.Text
    .Fields(4) = Combo2.Text
    .Fields(5) = DTPicker1.Value
    .Fields(6) = Text1.Text
    .Fields(7) = file
    .Update
End With
End Function

Private Sub Form_Unload(Cancel As Integer)
Login.Text5.Text = ""
Login.Text7.Text = ""
Login.Combo4.Text = ""
End Sub
