VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form ADODC 
   Caption         =   "Form1"
   ClientHeight    =   7275
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10620
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
   ScaleHeight     =   7275
   ScaleWidth      =   10620
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Append"
      Height          =   855
      Index           =   0
      Left            =   2640
      TabIndex        =   11
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Edit"
      Height          =   855
      Index           =   1
      Left            =   4320
      TabIndex        =   10
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete"
      Height          =   855
      Index           =   2
      Left            =   6000
      TabIndex        =   9
      Top             =   4560
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<<"
      Height          =   855
      Index           =   3
      Left            =   2640
      TabIndex        =   8
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   855
      Index           =   4
      Left            =   4320
      TabIndex        =   7
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   855
      Index           =   5
      Left            =   6000
      TabIndex        =   6
      Top             =   5400
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      DataField       =   "name"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Text            =   " "
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "class"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Text            =   " "
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      DataField       =   "percentage"
      DataSource      =   "Adodc1"
      Enabled         =   0   'False
      Height          =   495
      Left            =   5880
      TabIndex        =   0
      Text            =   " "
      Top             =   2880
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   1560
      Top             =   3600
      Visible         =   0   'False
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   255
      ForeColor       =   16777215
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Users\user\Documents\My Data Sources\mydb.mdb"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Users\user\Documents\My Data Sources\mydb.mdb"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "new"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Name"
      Height          =   375
      Left            =   3120
      TabIndex        =   5
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Class"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Percentage"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   2880
      Width           =   1665
   End
End
Attribute VB_Name = "ADODC"
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
        Adodc1.Recordset.AddNew
    Case 1
        Text1.Enabled = True
        Text2.Enabled = True
        Text3.Enabled = True
        Adodc1.Recordset.Update
    Case 2
        Adodc1.Recordset.Delete
        If Adodc1.Recordset.EOF Then
            MsgBox "last record"
            Adodc1.Recordset.MoveFirst
        Else
            Adodc1.Recordset.MoveNext
        End If
    Case 3
        If Adodc1.Recordset.BOF Then
            MsgBox "First Record"
            Adodc1.Recordset.MoveLast
        Else
            Adodc1.Recordset.MovePrevious
        End If
    Case 4
        If Adodc1.Recordset.EOF Then
            MsgBox "last record"
            Adodc1.Recordset.MoveFirst
        Else
            Adodc1.Recordset.MoveNext
        End If
    Case 5
        End
End Select
End Sub
