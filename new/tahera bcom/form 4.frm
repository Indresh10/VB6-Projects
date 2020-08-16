VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form sales 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Form4"
   ClientHeight    =   9255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16815
   LinkTopic       =   "Form4"
   ScaleHeight     =   9255
   ScaleWidth      =   16815
   Begin MSAdodcLib.Adodc Adodc3 
      Height          =   375
      Left            =   3000
      Top             =   840
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
      RecordSource    =   "sales"
      Caption         =   "Sales"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Print Bill"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   15
      Top             =   8280
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   1320
      Width           =   1575
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   615
      Left            =   10680
      Top             =   1800
      Visible         =   0   'False
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   1085
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
      RecordSource    =   "customer"
      Caption         =   "Fruits"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Preview Bill"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2160
      TabIndex        =   12
      Top             =   8280
      Width           =   2295
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   13560
      TabIndex        =   11
      Text            =   "Combo2"
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   13560
      TabIndex        =   7
      Top             =   3720
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   495
      Left            =   13560
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "form 4.frx":0000
      Height          =   4695
      Left            =   960
      TabIndex        =   5
      Top             =   2640
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8281
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
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
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "REMOVE"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5520
      Width           =   1500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000010&
      Caption         =   "ADD"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10800
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5520
      Width           =   1500
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   495
      Left            =   1560
      Top             =   8280
      Visible         =   0   'False
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
      RecordSource    =   "select * from bill"
      Caption         =   "Bill"
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
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   420
      Left            =   8715
      TabIndex        =   17
      Top             =   7560
      Width           =   255
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TOTAL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   6840
      TabIndex        =   16
      Top             =   7560
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BILL NO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   3480
      TabIndex        =   13
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF FRUITS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   10800
      TabIndex        =   10
      Top             =   3000
      Width           =   2610
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "QUANTITY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   10800
      TabIndex        =   9
      Top             =   3840
      Width           =   1560
   End
   Begin VB.Label s 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   10800
      TabIndex        =   8
      Top             =   4680
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NAME OF CUSTOMER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404080&
      Height          =   360
      Left            =   1320
      TabIndex        =   1
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SALES DETAILS"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   0
      Top             =   360
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0C0FF&
      FillStyle       =   6  'Cross
      Height          =   3615
      Left            =   10680
      Shape           =   4  'Rounded Rectangle
      Top             =   2760
      Width           =   4575
   End
End
Attribute VB_Name = "sales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim total As Single
Private Sub Command1_Click()
Adodc2.Refresh
If Adodc2.Recordset.RecordCount > 0 Then
    Adodc2.Recordset.MoveLast
    id = Adodc2.Recordset.Fields(0) + 1
Else
    id = 1
End If
Adodc2.Recordset.AddNew
Adodc2.Recordset.Fields(0) = id
Adodc2.Recordset.Fields(1) = Combo2.Text
Adodc2.Recordset.Fields(2) = Val(Text3.Text)
Adodc2.Recordset.Fields(3) = Val(Text4.Text)
Adodc2.Recordset.Fields(4) = Val(Text3.Text) * Val(Text4.Text)
Adodc2.Recordset.Update
Adodc2.Refresh
total = total + (Val(Text3.Text) * Val(Text4.Text))
Label7.Caption = total
DataGrid1.Refresh
Adodc1.Recordset.Find ("Fruit='" + Combo2.Text + "'")
Adodc1.Recordset.Fields(1) = Adodc1.Recordset.Fields(1) - Val(Text3.Text)
Adodc1.Recordset.Update
Adodc1.Refresh
Adodc2.Refresh
DataGrid1.Refresh
Combo2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Command2_Click()
m = InputBox("Enter The Item No.")
Adodc2.Recordset.MoveFirst
Adodc2.Recordset.Find ("Itemno=" & Val(m))
If Not Adodc2.Recordset.EOF Then
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("Fruit='" + Adodc2.Recordset.Fields(1) + "'")
Adodc1.Recordset.Update 1, Adodc1.Recordset.Fields(1) + Adodc2.Recordset.Fields(2)
Adodc1.Refresh
Adodc2.Recordset.Delete
Adodc2.Refresh
End If
Adodc2.Refresh
total = total - (Val(Text3.Text) * Val(Text4.Text))
Label7.Caption = total
DataGrid1.Refresh
DataGrid1.Refresh
End Sub

Private Sub Command3_Click()
Set DataReport1.DataSource = DataGrid1.DataSource
DataReport1.Sections("Section2").Controls("Label4").Caption = Text1.Text
DataReport1.Sections("Section2").Controls("Label5").Caption = Combo1.Text
DataReport1.Sections("Section3").Controls("Label12").Caption = Label7.Caption
DataReport1.PrintReport True
End Sub

Private Sub Command4_Click()
Set DataReport1.DataSource = DataGrid1.DataSource
DataReport1.Sections("Section2").Controls("Label4").Caption = Text1.Text
DataReport1.Sections("Section2").Controls("Label5").Caption = Combo1.Text
DataReport1.Sections("Section3").Controls("Label12").Caption = Label7.Caption
DataReport1.Show
End Sub

Private Sub Form_Load()
total = 0
Adodc1.Refresh
Combo1.Clear
While Not Adodc1.Recordset.EOF
    Combo1.AddItem Adodc1.Recordset.Fields(1)
    Adodc1.Recordset.MoveNext
Wend
Adodc1.RecordSource = "Fruit"
Adodc1.Refresh
Combo2.Clear
While Not Adodc1.Recordset.EOF
    Combo2.AddItem Adodc1.Recordset.Fields(0)
    Adodc1.Recordset.MoveNext
Wend
Adodc3.Refresh
If Not (Adodc3.Recordset.BOF And Adodc3.Recordset.EOF) Then Adodc3.Recordset.MoveLast
If Adodc3.Recordset.BOF Then
    Text1.Text = 1
Else
    Text1.Text = Adodc1.Recordset.Fields(0) + 1
End If
Label7.Caption = total
End Sub

Private Sub Text3_Validate(Cancel As Boolean)
Adodc1.Recordset.MoveFirst
Adodc1.Recordset.Find ("Fruit='" + Combo2.Text + "'")
If Not Adodc1.Recordset.EOF Then
    If Val(Text3.Text) > Adodc1.Recordset.Fields(1) Then
        MsgBox "Fruit not Available in that much quantity"
        Text3.SetFocus
    End If
End If
End Sub

