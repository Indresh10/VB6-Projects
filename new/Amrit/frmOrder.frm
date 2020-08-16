VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmOrder 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Order"
   ClientHeight    =   7575
   ClientLeft      =   4695
   ClientTop       =   1035
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   -1  'True
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FF00FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmOrder.frx":0000
   ScaleHeight     =   7575
   ScaleWidth      =   12225
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   615
      Left            =   2400
      Top             =   6840
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Items"
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   360
      Top             =   6840
      Visible         =   0   'False
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   1085
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=G:\restaurant\restaurant.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Select * from Orders "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmOrder.frx":3ABB8
      Height          =   3015
      Left            =   120
      TabIndex        =   23
      Top             =   3360
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   5318
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   16393
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   6720
      TabIndex        =   20
      Top             =   6600
      Width           =   2775
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   375
         Left            =   1200
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Total"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame fraAction 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Action"
      Height          =   2595
      Left            =   9960
      TabIndex        =   15
      Top             =   3540
      Width           =   2205
      Begin VB.CommandButton Command1 
         Caption         =   "&Print Bill and Reset"
         Height          =   705
         Left            =   360
         TabIndex        =   18
         Top             =   1080
         Width           =   1470
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Preview &Bill"
         Height          =   705
         Left            =   360
         TabIndex        =   3
         Top             =   345
         Width           =   1470
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Cancel"
         Height          =   585
         Left            =   360
         TabIndex        =   4
         Top             =   1830
         Width           =   1470
      End
   End
   Begin VB.Frame fraItems 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Items Information"
      Height          =   2250
      Left            =   5700
      TabIndex        =   11
      Top             =   870
      Width           =   6405
      Begin VB.CommandButton Command3 
         Caption         =   "&Remove from Cart"
         Height          =   345
         Left            =   3240
         TabIndex        =   24
         Top             =   1800
         Width           =   2430
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Add to Cart"
         Height          =   345
         Left            =   1320
         TabIndex        =   19
         Top             =   1800
         Width           =   1830
      End
      Begin VB.TextBox txtqty 
         Height          =   390
         Left            =   4875
         MaxLength       =   4
         TabIndex        =   2
         Top             =   1320
         Width           =   1245
      End
      Begin VB.TextBox txtPrice 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2700
         TabIndex        =   6
         Top             =   1305
         Width           =   1365
      End
      Begin VB.ComboBox cboItems 
         Height          =   390
         Left            =   2685
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   825
         Width           =   3465
      End
      Begin VB.ComboBox cbocategory 
         Height          =   390
         Left            =   2700
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   285
         Width           =   3450
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   345
         Left            =   4215
         TabIndex        =   16
         Top             =   1365
         Width           =   825
      End
      Begin VB.Label lblPrice 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   525
         Left            =   300
         TabIndex        =   14
         Top             =   1335
         Width           =   1335
      End
      Begin VB.Label lblAddItems 
         BackStyle       =   0  'Transparent
         Caption         =   "Items"
         Height          =   525
         Left            =   330
         TabIndex        =   13
         Top             =   855
         Width           =   1335
      End
      Begin VB.Label lblCategory 
         BackStyle       =   0  'Transparent
         Caption         =   "Category"
         Height          =   525
         Left            =   330
         TabIndex        =   12
         Top             =   345
         Width           =   1335
      End
   End
   Begin VB.Frame fraAddOrder 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Add Order"
      Height          =   1440
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   5565
      Begin VB.TextBox Text2 
         Height          =   390
         Left            =   2880
         MaxLength       =   8
         TabIndex        =   17
         Top             =   840
         Width           =   2385
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   390
         Left            =   2880
         TabIndex        =   5
         Top             =   300
         Width           =   2385
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Customer"
         Height          =   465
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   855
         Width           =   1305
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         Caption         =   "Order No"
         Height          =   465
         Index           =   0
         Left            =   210
         TabIndex        =   9
         Top             =   390
         Width           =   1305
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H000000FF&
      Caption         =   "Order Information"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   615
      Left            =   3660
      TabIndex        =   7
      Top             =   60
      Width           =   4935
   End
End
Attribute VB_Name = "frmOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Single
Private Sub cbocategory_Click()
Adodc2.RecordSource = "Select * from Items where cat_name='" + cbocategory.Text + "'"
Adodc2.Refresh
cboItems.Clear
cbocategory.AddItem " "
While Not Adodc2.Recordset.EOF
    If Not IsNull(Adodc2.Recordset.Fields("Item_Name")) Then cboItems.AddItem Adodc2.Recordset.Fields("Item_Name")
    Adodc2.Recordset.MoveNext
Wend
End Sub

Private Sub cboItems_Click()
Adodc2.RecordSource = "Select * from Items where Item_Name='" + cboItems.Text + "'"
Adodc2.Refresh
txtPrice.Text = Adodc2.Recordset.Fields("Rate")
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Adodc1.Refresh
DRBill.Sections("PageHeader").Controls("Label6").Caption = Text1.Text
DRBill.Sections("PageHeader").Controls("Label7").Caption = Text2.Text
DRBill.Sections("PageFooter").Controls("Label5").Caption = Label4.Caption
DRBill.Show
End Sub

Private Sub Command1_Click()
Adodc1.Refresh
Adodc2.RecordSource = "Select * from Customer"
Adodc2.Refresh
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = Text1.Text
Adodc2.Recordset.Fields(1) = Text2.Text
Adodc2.Recordset.Fields(2) = Label4.Caption
Adodc2.Recordset.Update
Adodc2.Refresh
DRBill.Sections("PageHeader").Controls("Label6").Caption = Text1.Text
DRBill.Sections("PageHeader").Controls("Label7").Caption = Text2.Text
DRBill.Sections("PageFooter").Controls("Label5").Caption = Label4.Caption
DRBill.PrintReport True
cbocategory.ListIndex = 0
cbocategory.ListIndex = 0
txtPrice.Text = ""
txtqty.Text = ""
Text2.Text = ""
Call Form_Load
End Sub

Private Sub Command2_Click()
If Val(txtqty) > 0 Then
Adodc1.Refresh
If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveLast
    id = Adodc1.Recordset.Fields(0) + 1
Else
    id = 1
End If
Adodc1.Recordset.AddNew
Adodc1.Recordset.Fields(0) = id
Adodc1.Recordset.Fields(1) = cboItems.Text
Adodc1.Recordset.Fields(2) = txtPrice.Text
Adodc1.Recordset.Fields(3) = txtqty.Text
Adodc1.Recordset.Fields(4) = txtPrice.Text * txtqty.Text
total = total + (txtPrice.Text * txtqty.Text)
Adodc1.Recordset.Update
DataGrid1.Refresh
Label4.Caption = total
cbocategory.ListIndex = 0
cbocategory.ListIndex = 0
txtPrice.Text = ""
txtqty.Text = ""
Else
    MsgBox "Quantity Must be Greater than 0"
End If
End Sub

Private Sub Command3_Click()
m = Val(InputBox("Enter the Item no"))
Adodc1.Recordset.Find "Item_No=" & m
If Not Adodc1.Recordset.EOF Then
total = total - Adodc1.Recordset.Fields("Subtotal")
Label4.Caption = total
Adodc1.Recordset.Delete
End If
Call refr
Call refr
End Sub

Private Sub Form_Load()
total = 0
If Adodc1.Recordset.RecordCount > 0 Then
    Adodc1.Recordset.MoveFirst
    While Not Adodc1.Recordset.EOF
        Adodc1.Recordset.Delete
        Adodc1.Recordset.MoveNext
    Wend
End If
Adodc1.Refresh
DataGrid1.Refresh
Adodc2.RecordSource = "Select * from Customer"
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveLast
    id = Adodc2.Recordset.Fields(0) + 1
Else
    id = 1
End If
Text1.Text = id
Adodc2.RecordSource = "Select distinct cat_name from Items"
Adodc2.Refresh
cbocategory.AddItem " "
While Not Adodc2.Recordset.EOF
    cbocategory.AddItem Adodc2.Recordset.Fields(0)
    Adodc2.Recordset.MoveNext
Wend
Call refr
Call refr
Label4.Caption = total
End Sub

Private Sub txtqty_KeyPress(KeyAscii As Integer)
If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or (KeyAscii = 8) Or (KeyAscii = 46)) Then
            KeyAscii = 0
            MsgBox "Please Enter Numeric Value "
End If
End Sub

Public Function refr()
Adodc1.Refresh
DataGrid1.Refresh
End Function
