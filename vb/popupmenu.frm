VERSION 5.00
Begin VB.Form popupmenu 
   Caption         =   "Form1"
   ClientHeight    =   6495
   ClientLeft      =   225
   ClientTop       =   570
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   10440
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Shape Shape1 
      Height          =   4335
      Left            =   600
      Top             =   1080
      Width           =   8775
   End
   Begin VB.Menu shapemenu 
      Caption         =   "shape"
      Visible         =   0   'False
      Begin VB.Menu shape 
         Caption         =   "rectangle"
         Index           =   0
      End
      Begin VB.Menu shape 
         Caption         =   "square"
         Index           =   1
      End
      Begin VB.Menu shape 
         Caption         =   "oval"
         Index           =   2
      End
      Begin VB.Menu shape 
         Caption         =   "circle"
         Index           =   3
      End
      Begin VB.Menu shape 
         Caption         =   "rounded rectangle"
         Index           =   4
      End
      Begin VB.Menu shape 
         Caption         =   "rounded square"
         Index           =   5
      End
   End
End
Attribute VB_Name = "popupmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then popupmenu shapemenu
End Sub

Private Sub shape_Click(Index As Integer)
Shape1.shape = Index
End Sub
