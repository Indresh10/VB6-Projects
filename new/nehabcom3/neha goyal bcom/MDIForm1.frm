VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu bif 
      Caption         =   "BOOK Issue Form"
   End
   Begin VB.Menu br 
      Caption         =   "Book Record"
   End
   Begin VB.Menu brf 
      Caption         =   "Book Return Form"
   End
   Begin VB.Menu sr 
      Caption         =   "Student Record"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub bif_Click()
Form2.Show
End Sub

Private Sub br_Click()
Form3.Show
End Sub

Private Sub brf_Click()
Form4.Show
End Sub

Private Sub sr_Click()
Form5.Show
End Sub
