VERSION 5.00
Begin VB.Form data 
   Caption         =   "Form1"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   15
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   12330
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   855
      Index           =   5
      Left            =   6720
      TabIndex        =   11
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   855
      Index           =   4
      Left            =   5040
      TabIndex        =   10
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   855
      Index           =   3
      Left            =   3360
      TabIndex        =   9
      Top             =   5520
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   855
      Index           =   2
      Left            =   6720
      TabIndex        =   8
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   855
      Index           =   1
      Left            =   5040
      TabIndex        =   7
      Top             =   4680
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Append"
      Height          =   855
      Index           =   0
      Left            =   3360
      TabIndex        =   6
      Top             =   4680
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      DataField       =   "percentage"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   5
      Text            =   " "
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "class"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   4
      Text            =   " "
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Data1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   6240
      TabIndex        =   3
      Text            =   " "
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Users\user\Documents\My Data Sources\mydb.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "new"
      Top             =   3720
      Width           =   5895
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Percentage"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   2760
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Class"
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   1080
      Width           =   855
   End
End
Attribute VB_Name = "data"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
Text1.Enabled = False
Text2.Enabled = False
Text3.Enabled = False
Select Case Index
    Case 0
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Data1.Recordset.AddNew
    Case 1
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Data1.Recordset.Edit
    Case 2
        Data1.Recordset.Delete
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveFirst
        Else
            Data1.Recordset.MoveNext
        End If
    Case 3
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveLast
        Else
            Data1.Recordset.MovePrevious
        End If
    Case 4
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveFirst
        Else
            Data1.Recordset.MoveNext
        End If
    Case 5
        End
End Select
End Sub


Private Sub Form_Load()
Data1.Visible = False
End Sub
