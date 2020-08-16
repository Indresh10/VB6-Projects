VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form progressbar 
   Caption         =   "Form2"
   ClientHeight    =   4440
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9900
   LinkTopic       =   "Form2"
   ScaleHeight     =   4440
   ScaleWidth      =   9900
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3480
      Top             =   2400
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   1800
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   1508
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "progressbar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x
Private Sub Form_Load()
x = 0
End Sub

Private Sub Timer1_Timer()
x = x + 10
If x > 100 Then
m = MsgBox("Succesfully installed")
End
End If
ProgressBar1.Value = x
End Sub
