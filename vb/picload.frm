VERSION 5.00
Begin VB.Form picload 
   Caption         =   "Form1"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6795
   LinkTopic       =   "Form1"
   ScaleHeight     =   5670
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2655
      Left            =   840
      Picture         =   "picload.frx":0000
      ScaleHeight     =   2595
      ScaleWidth      =   4635
      TabIndex        =   3
      Top             =   2760
      Width           =   4695
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   4560
      Pattern         =   "*.jpg;*.bmp;*.png"
      System          =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
End
Attribute VB_Name = "picload"
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
fn = File1.Path + "\" + File1.FileName
Picture1.Picture = LoadPicture(fn)
End Sub

