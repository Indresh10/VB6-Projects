VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7185
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   7185
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   3495
      Left            =   240
      ScaleHeight     =   3435
      ScaleWidth      =   6315
      TabIndex        =   3
      Top             =   3240
      Width           =   6375
   End
   Begin VB.FileListBox File1 
      Height          =   2340
      Left            =   3480
      Pattern         =   "*.bmp;*.jpg;*.png"
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   1710
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2895
   End
   Begin VB.DriveListBox Drive1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
    File1.Path = Dir1
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1
End Sub

Private Sub File1_Click()
    fn = File1.Path + "/" + File1.FileName
    Picture1.Picture = LoadPicture(fn)
End Sub
