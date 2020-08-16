VERSION 5.00
Begin VB.Form menu 
   Caption         =   "Form1"
   ClientHeight    =   6885
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   ScaleHeight     =   6885
   ScaleWidth      =   12360
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Shape Shape1 
      Height          =   3975
      Left            =   1560
      Top             =   480
      Width           =   9255
   End
   Begin VB.Menu shapemenu 
      Caption         =   "shape"
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
         Caption         =   "rounded square "
         Index           =   5
      End
   End
   Begin VB.Menu fills_menu 
      Caption         =   "fillstyle"
      Begin VB.Menu fillstyle 
         Caption         =   "solid"
         Index           =   0
      End
      Begin VB.Menu fillstyle 
         Caption         =   "transparent"
         Index           =   1
      End
      Begin VB.Menu fillstyle 
         Caption         =   "horizontal line"
         Index           =   2
      End
      Begin VB.Menu fillstyle 
         Caption         =   "vertical line"
         Index           =   3
      End
      Begin VB.Menu fillstyle 
         Caption         =   "upward diagonal"
         Index           =   4
      End
      Begin VB.Menu fillstyle 
         Caption         =   "downward diagonal"
         Index           =   5
      End
      Begin VB.Menu fillstyle 
         Caption         =   "cross"
         Index           =   6
      End
      Begin VB.Menu fillstyle 
         Caption         =   "diagonal cross"
         Index           =   7
      End
   End
End
Attribute VB_Name = "menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub fillstyle_Click(Index As Integer)
Shape1.fillstyle = Index
End Sub

Private Sub shape_Click(Index As Integer)
Shape1.shape = Index
End Sub
