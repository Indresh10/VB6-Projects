VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form purchase 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form5"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12690
   LinkTopic       =   "Form5"
   ScaleHeight     =   9525
   ScaleWidth      =   12690
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   330
      Left            =   1200
      Top             =   7680
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Fruit"
      Caption         =   "Adodc3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   495
      Left            =   7200
      TabIndex        =   19
      Top             =   5520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      _Version        =   393216
      Format          =   134479873
      CurrentDate     =   43853
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   4080
      Top             =   1440
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   582
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "supplier"
      Caption         =   "Supplier"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text6 
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      MaxLength       =   10
      TabIndex        =   18
      Top             =   3600
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   7200
      TabIndex        =   17
      Top             =   7080
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   7200
      TabIndex        =   15
      Text            =   "Combo1"
      Top             =   1920
      Width           =   3375
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000A&
      Caption         =   "EXIT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8760
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000A&
      Caption         =   "SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8760
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   495
      Left            =   7200
      TabIndex        =   12
      Top             =   7800
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   7200
      TabIndex        =   10
      Top             =   6360
      Width           =   3255
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cheque"
      Height          =   495
      Left            =   8880
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Cash"
      Height          =   495
      Left            =   7320
      TabIndex        =   6
      Top             =   4560
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Width           =   3375
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   375
      Left            =   3720
      Top             =   8760
      Visible         =   0   'False
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   661
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Indresh Hemani\Desktop\new\tahera bcom\Fruit.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "purchase"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   16
      Top             =   7200
      Width           =   1455
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   7920
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE FRUIT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3000
      TabIndex        =   9
      Top             =   6360
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "DATE OF PURCHASE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3000
      TabIndex        =   8
      Top             =   5520
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "PAYMENT"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   5
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "CONTACT NO."
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   3720
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "ADDRESS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF THE SUPPLIER"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3000
      TabIndex        =   1
      Top             =   2040
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PURCHASE DETAILS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   525
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   4860
   End
End
Attribute VB_Name = "purchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo1_Change()
If Not Combo1.Text = "" Then
If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("sname='" + Combo1.Text + "'")
Text2.Text = Adodc1.Recordset.Fields(2)
Text6.Text = Adodc1.Recordset.Fields(3)
End If
End If
End Sub

Private Sub Combo1_Click()
If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("sname='" + Combo1.Text + "'")
Text2.Text = Adodc1.Recordset.Fields(2)
Text6.Text = Adodc1.Recordset.Fields(3)
End If
End Sub

Private Sub Command1_Click()
Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Adodc1.Recordset.Fields(0)
Adodc2.Recordset.Fields(1) = Combo1.Text
If Option1.Value = True Then
    Adodc2.Recordset.Fields(2) = Option1.Caption
ElseIf Option2.Value = True Then
    Adodc2.Recordset.Fields(2) = Option2.Caption
End If
Adodc2.Recordset.Fields(3) = DTPicker1.Value
Adodc2.Recordset.Fields(4) = UCase(Text4.Text)
Adodc2.Recordset.Fields(5) = Val(Text1.Text)
Adodc2.Recordset.Fields(6) = Val(Text5.Text)
Adodc2.Recordset.Update
Adodc2.Refresh
Adodc1.RecordSource = "Fruit"
Adodc1.Refresh
If Not (Adodc1.Recordset.EOF And Adodc1.Recordset.BOF) Then Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("Fruit='" + UCase(Text4.Text) + "'")
If Adodc1.Recordset.EOF Then
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields(0) = UCase(Text4.Text)
    Adodc1.Recordset.Fields(1) = Val(Text5.Text)
Else
    Adodc1.Recordset.Fields(1) = Adodc1.Recordset.Fields(1) + Val(Text5.Text)
End If
Adodc1.Recordset.Update
Adodc1.RecordSource = "Supplier"
Adodc1.Refresh
MsgBox "saved successfully"
Combo1.Text = ""
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Option1.Value = False
Option2.Value = False
DTPicker1.Value = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Combo1.Clear
Adodc1.Refresh
Adodc1.Recordset.MoveFirst
While Not Adodc1.Recordset.EOF
    Combo1.AddItem Adodc1.Recordset.Fields(1)
    Adodc1.Recordset.MoveNext
Wend
DTPicker1.Value = Format(Now, "dd/mm/yyyy")
End Sub
