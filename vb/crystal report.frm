VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#4.6#0"; "crystl32.ocx"
Begin VB.Form cry_rep 
   Caption         =   "Form1"
   ClientHeight    =   6030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   6030
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
      Height          =   495
      Left            =   3720
      TabIndex        =   0
      Top             =   2760
      Width           =   1215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   1440
      Top             =   1920
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   262150
      ReportFileName  =   "D:\vb\cr1.rpt"
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
   End
End
Attribute VB_Name = "cry_rep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
CrystalReport1.PrintReport
End Sub
