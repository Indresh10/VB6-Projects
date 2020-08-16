VERSION 5.00
Begin VB.Form FrmTransfer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Member Transfer"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   Icon            =   "Member_Transfer_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   7140
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   5640
      TabIndex        =   13
      Top             =   5415
      Width           =   1065
   End
   Begin VB.CommandButton CmdTransfer 
      Caption         =   "Transfer"
      Default         =   -1  'True
      Height          =   435
      Left            =   375
      TabIndex        =   12
      Top             =   5430
      Width           =   1065
   End
   Begin VB.Frame FremTo 
      Caption         =   "Transfer To"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   3645
      TabIndex        =   6
      Top             =   105
      Width           =   3375
      Begin VB.ComboBox CmbClassTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Member_Transfer_Form.frx":030A
         Left            =   1380
         List            =   "Member_Transfer_Form.frx":0326
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   405
         Width           =   1215
      End
      Begin VB.ComboBox CmbClassYearTo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Member_Transfer_Form.frx":0359
         Left            =   1380
         List            =   "Member_Transfer_Form.frx":0378
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   870
         Width           =   1215
      End
      Begin VB.ListBox LstTo 
         Height          =   3570
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   11
         Top             =   1440
         Width           =   3240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Class :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   750
         TabIndex        =   7
         Top             =   465
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   750
         TabIndex        =   9
         Top             =   930
         Width           =   525
      End
   End
   Begin VB.Frame FremFrom 
      Caption         =   "Transfer From"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5100
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   3375
      Begin VB.ListBox LstFrom 
         Height          =   3570
         Left            =   60
         Sorted          =   -1  'True
         TabIndex        =   5
         Top             =   1440
         Width           =   3240
      End
      Begin VB.ComboBox CmbClassYearFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Member_Transfer_Form.frx":03AC
         Left            =   1320
         List            =   "Member_Transfer_Form.frx":03CB
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   870
         Width           =   1215
      End
      Begin VB.ComboBox CmbClassFrom 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Member_Transfer_Form.frx":03FF
         Left            =   1320
         List            =   "Member_Transfer_Form.frx":041B
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   405
         Width           =   1215
      End
      Begin VB.Label LblYear 
         AutoSize        =   -1  'True
         Caption         =   "Year :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   690
         TabIndex        =   3
         Top             =   930
         Width           =   525
      End
      Begin VB.Label LblClass 
         AutoSize        =   -1  'True
         Caption         =   "Class :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   690
         TabIndex        =   1
         Top             =   465
         Width           =   600
      End
   End
End
Attribute VB_Name = "FrmTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset

Dim i As Integer, cnt As Integer
Dim Qry As String


Private Sub CmbClassFrom_Click()
    Call fillYear(CmbClassFrom, CmbClassYearFrom)
    CmbClassYearFrom.Text = CmbClassYearFrom.List(0)
End Sub

Private Sub CmbClassTo_Click()
    Call fillYear(CmbClassTo, CmbClassYearTo)
    CmbClassYearTo.Text = CmbClassYearTo.List(0)
End Sub

Private Sub CmbClassYearFrom_Click()
    
    Set rs = New Recordset
    rs.Open "SELECT * FROM Mbr_Mast WHERE [Crs]='" & CmbClassFrom & _
            "' AND [Yer]='" & CmbClassYearFrom & "'", conn, adOpenStatic, adLockReadOnly
    
    LstFrom.Clear
    If rs.RecordCount > 0 Then

        Do While Not rs.EOF
            LstFrom.AddItem rs.Fields(0) & " " & _
                    rs.Fields(1) & " " & rs.Fields(2) & " " & rs.Fields(3)
            
            rs.MoveNext
        Loop
    
    End If
End Sub

Private Sub CmbClassYearTo_Click()
    Set rs1 = New Recordset
    rs1.Open "SELECT * FROM Mbr_Mast WHERE [Crs]='" & CmbClassTo.Text & _
            "' AND [Yer]='" & CmbClassYearTo.Text & "' ORDER BY Code", conn, adOpenStatic, adLockReadOnly
            
    LstTo.Clear
    If rs1.RecordCount > 0 Then

        Do While Not rs1.EOF
            LstTo.AddItem rs1.Fields(0) & " " & _
                    rs1.Fields(1) & " " & rs1.Fields(2) & " " & rs1.Fields(3)
            
            rs1.MoveNext
        Loop

    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
    FrmMember.Show
End Sub

Private Sub CmdTransfer_Click()
    'WHEN TRANSFER TO SAME CLASS & YEAR
    If CmbClassFrom.Text = CmbClassTo.Text And CmbClassYearFrom.Text = CmbClassYearTo.Text Then
        MsgBox "Member can not transfer to same class and year", vbInformation, "Member Transfer"
        Exit Sub
    End If
            
    If rs1.RecordCount = 0 Then
        Qry = "UPDATE Mbr_Mast SET [Crs]='" & CmbClassTo.Text & _
            "',[Yer]='" & CmbClassYearTo.Text & "' WHERE [Crs]='" & _
            CmbClassFrom.Text & "' AND [Yer]='" & CmbClassYearFrom.Text & "'"
        
        conn.Execute Qry
        
        MsgBox "Member transmitted successfully.", vbInformation, "Member Transfer"
        
        LstFrom.Clear
        Call CmbClassYearTo_Click   'TO RETRIVE UPDATED DATA
    Else
        MsgBox "Destination Class is not empty.", vbInformation, "Member Transfer"
    End If
End Sub

Private Sub Form_Load()
    
    CmbClassFrom.Text = CmbClassFrom.List(0)
    CmbClassTo.Text = CmbClassTo.List(0)
    
End Sub


'========================================================
'FILL YEAR COMBO BOX
Public Sub fillYear(c As Control, y As Control)
        
    y.Clear
    If c.Text = "BBA" Or c.Text = "BCOM" Then
        
        y.AddItem "FY"
        y.AddItem "SY"
        y.AddItem "TY"
        
    ElseIf c.Text = "PGDCA" Or c.Text = "DCS" Then
        For i = 1 To 2
            y.AddItem "SEM" & i
        Next
    Else
        For i = 1 To 6
            y.AddItem "SEM" & i
        Next
    End If

End Sub

