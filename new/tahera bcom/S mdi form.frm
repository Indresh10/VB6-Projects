VERSION 5.00
Begin VB.MDIForm home 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   3030
   ClientLeft      =   225
   ClientTop       =   1170
   ClientWidth     =   4560
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu SD 
      Caption         =   "supplier detail"
   End
   Begin VB.Menu SAL 
      Caption         =   "SALES"
   End
   Begin VB.Menu PUR 
      Caption         =   "PURCHASE"
   End
   Begin VB.Menu cd 
      Caption         =   "Customer detail"
   End
End
Attribute VB_Name = "home"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cd_Click()
customer.Show
End Sub

Private Sub PUR_Click()
purchase.Show
End Sub

Private Sub SAL_Click()
sales.Show
End Sub

Private Sub SD_Click()
supplier.Show
End Sub
