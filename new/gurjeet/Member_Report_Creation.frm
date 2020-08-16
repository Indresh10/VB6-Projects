VERSION 5.00
Begin VB.Form FrmRpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Member Report Creation"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5310
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   3780
      TabIndex        =   6
      Top             =   2130
      Width           =   1140
   End
   Begin VB.CommandButton CmdCreate 
      Caption         =   "&Create"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   375
      TabIndex        =   5
      Top             =   2130
      Width           =   1140
   End
   Begin VB.Frame FremClass 
      Caption         =   "Choose Class && Year"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   105
      TabIndex        =   0
      Top             =   75
      Width           =   5100
      Begin VB.ComboBox CmbClass 
         Height          =   360
         ItemData        =   "Member_Report_Creation.frx":0000
         Left            =   960
         List            =   "Member_Report_Creation.frx":001C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   810
         Width           =   1215
      End
      Begin VB.ComboBox CmbClassYear 
         Height          =   360
         ItemData        =   "Member_Report_Creation.frx":004F
         Left            =   3585
         List            =   "Member_Report_Creation.frx":006E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label LblClass 
         AutoSize        =   -1  'True
         Caption         =   "Class :"
         Height          =   240
         Left            =   240
         TabIndex        =   1
         Top             =   870
         Width           =   600
      End
      Begin VB.Label LblYear 
         AutoSize        =   -1  'True
         Caption         =   "Year :"
         Height          =   240
         Left            =   3015
         TabIndex        =   3
         Top             =   870
         Width           =   525
      End
   End
End
Attribute VB_Name = "FrmRpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rs As New ADODB.Recordset
Dim FL As String


Private Sub CmbClass_Click()
    Call fillYear(Me) 'SELECT YEAR
    CmbClassYear.Text = CmbClassYear.List(0)
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub CmdCreate_Click()
    If Report = "M" Then
    'MEMBER REPORT
        
        Set rs = New Recordset
        rs.Open "SELECT Code,surname,member,father,Join_Dt,Cnt_No FROM Mbr_Mast WHERE [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "' ORDER BY Code", conn, adOpenStatic, adLockReadOnly
        
        'WHEN NO RECORD EXIST
        If rs.RecordCount = 0 Then
            rs.Close
            MsgBox "No record is found.", vbInformation, "Member Report"
            Exit Sub
        End If
        
        'CREATE REPORT
        'OPEN FILE
        FL = "Member_" & Format(Date, "dd-mm-yyyy")
        Open App.Path & "\Reports\" & FL & ".txt" For Output As #1
        
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "--------------------------- M E M B E R  R E P O R T ---------------------------"
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, "Class : " & CmbClass.Text & Space(50) & "Date : " & Format(Date, "dd-mm-yyyy")
        Print #1, "Year  : " & CmbClassYear.Text
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, " CODE   NAME                                            JOIN DATE    CONTACT NO."
        Print #1, "--------------------------------------------------------------------------------"
        
        rs.MoveFirst
        Do While Not rs.EOF
            Print #1, " " & rs.Fields(0) & " " & _
                rs.Fields(1) & " " & rs.Fields(2) & " " & rs.Fields(3) & _
                Space(50 - (Len(rs.Fields(1)) + Len(rs.Fields(2)) + Len(rs.Fields(3)) + 4)) & _
                Space(10 - Len(Format(rs.Fields(4), "dd-mm-yyyy"))) & Format(rs.Fields(4), "dd-mm-yyyy") & _
                Space(14 - Len(rs.Fields(5))) & rs.Fields(5)
            Print #1, ""
            rs.MoveNext
        Loop
        
        Close #1
        MsgBox FL & ".txt created successfully.", vbInformation, "Member Report"

        Shell App.Path & "\Reports\wordpad.exe " & App.Path & "\Reports\" & FL & ".txt", vbMaximizedFocus
    Else
    'ISSUE REPORT
        
        Set rs = New Recordset
        rs.Open "SELECT * FROM Issue_Mast WHERE [Crs]='" & CmbClass.Text & "' AND [Yer]='" & CmbClassYear.Text & "' ORDER BY Mbr_No", conn, adOpenStatic, adLockReadOnly
        
        'WHEN NO RECORD EXIST
        If rs.RecordCount = 0 Then
            rs.Close
            MsgBox "No record is found.", vbInformation, "Member Report"
            Exit Sub
        End If
        
        'CREATE REPORT
        'OPEN FILE
        FL = "Issue_" & Format(Date, "dd-mm-yyyy")
        Open App.Path & "\Reports\" & FL & ".txt" For Output As #1
        
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, "---------------------------- I S S U E  R E P O R T ----------------------------"
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, ""
        Print #1, "Class : " & CmbClass.Text & Space(50) & "Date : " & Format(Date, "dd-mm-yyyy")
        Print #1, "Year  : " & CmbClassYear.Text
        Print #1, ""
        Print #1, "--------------------------------------------------------------------------------"
        Print #1, " MEMBER CODE    BOOK CODE              ISSUED DATE             LAST SUBMIT DATE "
        Print #1, "--------------------------------------------------------------------------------"
        
        rs.MoveFirst
        Do While Not rs.EOF
            Print #1, " " & rs.Fields(0) & Space(15 - Len(rs.Fields(0))) & _
                    rs.Fields(3) & Space(23 - Len(rs.Fields(3))) & _
                    Format(rs.Fields(4), "dd-mmm-yyyy") & Space(24 - Len(Format(rs.Fields(4), "dd-mm-yyyy"))) & _
                    Format(rs.Fields(5), "dd-mmm-yyyy")

            Print #1, ""
            rs.MoveNext
        Loop
        
        Close #1
        MsgBox FL & ".txt created successfully.", vbInformation, "Member Report"
        
        Shell App.Path & "\Reports\wordpad.exe " & App.Path & "\Reports\" & FL & ".txt", vbMaximizedFocus
    End If
    
End Sub

Private Sub Form_Load()
    CmbClass.Text = CmbClass.List(0)
    
    If Report = "M" Then
    'MEMBER REPORT
        Me.Caption = "Member Report Creation"
    Else
    'ISSUE REPORT
        Me.Caption = "Issued/Submit Report Creation"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Forms.Count = 2 Then
        MDIFrm.Pct1.Visible = True
    End If
End Sub
