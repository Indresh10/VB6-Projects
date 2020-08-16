Attribute VB_Name = "Member"
Public cnt As Control
Public Class, Yer As String
Public Report As String

'SET CONTROLS ENABLE OR DISABLE
Public Sub controlEnable(Frm As Form, stat As Boolean)
    For Each cnt In Frm.Controls
        If TypeOf cnt Is TextBox Then
            cnt.Locked = Not stat
        End If
        
        If TypeOf cnt Is OptionButton Then
            cnt.Enabled = stat
        End If
    Next cnt
    
    Frm.CmbDay.Locked = Not stat
    Frm.CmbMonth.Locked = Not stat
    Frm.CmbYear.Locked = Not stat
End Sub

'SET DEFAULT CONTROLS
Public Sub clearControl(Frm As Form)
    For Each cnt In Frm.Controls
        If TypeOf cnt Is TextBox Then
            cnt.Text = ""
        End If
    Next cnt
    Frm.TxtFee.Text = 100
    Frm.CmbDay.Text = Day(Date)
    Frm.CmbMonth.Text = Month(Date)
    Frm.CmbYear.Text = Year(Date)
    Frm.CmbSearch.Text = Frm.CmbSearch.List(0)
    Frm.OptMale.Value = True
End Sub

'RETRIVE DATA OF MEMBER
Public Sub memberData(Frm As Form, ByVal rs As Recordset)
    With Frm
        .CmdSave.Enabled = False
        'RETRIVE RECORD
        .TxtCode.Text = rs.Fields(0)
        .TxtSurname.Text = rs.Fields(1)
        .TxtFirst.Text = rs.Fields(2)
        .TxtLast.Text = rs.Fields(3)
        .CmbDay.Text = Day(rs.Fields(4))
        .CmbMonth.Text = Month(rs.Fields(4))
        .CmbYear.Text = Year(rs.Fields(4))
        .TxtAddress.Text = rs.Fields(5)
        .TxtCity.Text = rs.Fields(6)
        .TxtFee.Text = rs.Fields(11)
        .TxtContact.Text = rs.Fields(9)
        If rs.Fields(10) = "M" Then
            .OptMale.Value = True
        Else
            .OptFemale.Value = True
        End If
    End With
End Sub

'======================================================
'FILL FLEX GRID OF MEMBER
Public Sub fillMbrGrid(Frm As Form, ByVal cr As String, yr As String, sr As String)
    Dim rs_tmp As New Recordset
    Dim r As Integer, Qry As String
    
    Set rs_tmp = New Recordset
    
    Qry = "select * from Mbr_Mast where Crs='" & cr & "' and Yer='" & yr & "' order by " & sr
    rs_tmp.Open Qry, conn, adOpenStatic
    
    With Frm
        'SET TITLE OF MSFlexGrid
        .MsfgSearch.Cols = 8
        .MsfgSearch.Rows = rs_tmp.RecordCount + 1

        If rs_tmp.RecordCount > 0 Then
            rs_tmp.MoveFirst
            For r = 1 To rs_tmp.RecordCount
                .MsfgSearch.TextMatrix(r, 0) = r
                .MsfgSearch.TextMatrix(r, 1) = rs_tmp.Fields(0)
                .MsfgSearch.TextMatrix(r, 2) = rs_tmp.Fields(1) & " " & rs_tmp.Fields(2) & " " & rs_tmp.Fields(3)
                .MsfgSearch.TextMatrix(r, 3) = Format(rs_tmp.Fields(4), "dd-mm-yyyy")
                .MsfgSearch.TextMatrix(r, 4) = rs_tmp.Fields(6)
                .MsfgSearch.TextMatrix(r, 5) = rs_tmp.Fields(9)
                .MsfgSearch.TextMatrix(r, 6) = rs_tmp.Fields(10)
                .MsfgSearch.TextMatrix(r, 7) = rs_tmp.Fields(12)
                rs_tmp.MoveNext
            Next
        End If
        
    End With
End Sub

'=======================================================
'ALPHA ALLOW
Public Function alpha(Key As Integer)
    
    If Key = 8 Then
        alpha = 8
    ElseIf (Key >= Asc("A") And Key <= Asc("Z")) Or _
        (Key >= Asc("a") And Key <= Asc("z")) Then
            
        alpha = Asc(UCase(Chr(Key)))
    Else
        alpha = 0
    End If
        
End Function

'========================================================
Public Sub fillYear(Frm As Form)
    Dim i As Integer
    
    With Frm
        
        .CmbClassYear.Clear
        If .CmbClass.Text = "BBA" Or .CmbClass = "BCOM" Then
        
            .CmbClassYear.AddItem "FY"
            .CmbClassYear.AddItem "SY"
            .CmbClassYear.AddItem "TY"
        
        ElseIf .CmbClass.Text = "PGDCA" Or .CmbClass.Text = "DCS" Then
            For i = 1 To 2
                .CmbClassYear.AddItem "SEM" & i
            Next
        Else
            For i = 1 To 6
                .CmbClassYear.AddItem "SEM" & i
            Next
        End If
        
    End With

End Sub

'==========================================================
'FILL GRID WHEN SEARCHING
Public Sub fillFlex(msf As MSFlexGrid, rs As Recordset)
    Dim r As Integer

    msf.Cols = rs.Fields.Count + 1
    msf.Rows = rs.RecordCount + 1
    
    For r = 1 To rs.RecordCount
        msf.TextMatrix(r, 0) = r
        For c = 1 To rs.Fields.Count
            If IsDate(rs.Fields(c - 1)) Then
                msf.TextMatrix(r, c) = Format(rs.Fields(c - 1), "dd-mm-yyyy")
            Else
                msf.TextMatrix(r, c) = rs.Fields(c - 1)
            End If
        Next
        rs.MoveNext
    Next
    
End Sub

'=============================================================
Public Function daysOfMonth(m As Integer, y As Integer) As Integer
    Select Case m
        Case 1, 3, 5, 7, 8, 10, 12
                daysOfMonth = 31
        Case 2
                If y Mod 4 = 0 Then
                    daysOfMonth = 29
                Else
                    daysOfMonth = 28
                End If
        Case Else
                daysOfMonth = 30
    End Select
End Function

'RETURN MONTH FOR BACKUP
Public Function Month_Nm() As String
    If Month(Date) = 1 Then
        Month_Nm = "JANUARY"
    ElseIf Month(Date) = 2 Then
        Month_Nm = "FEBRUARY"
    ElseIf Month(Date) = 3 Then
        Month_Nm = "MARCH"
    ElseIf Month(Date) = 4 Then
        Month_Nm = "APRIL"
    ElseIf Month(Date) = 5 Then
        Month_Nm = "MAY"
    ElseIf Month(Date) = 6 Then
        Month_Nm = "JUN"
    ElseIf Month(Date) = 7 Then
        Month_Nm = "JULY"
    ElseIf Month(Date) = 8 Then
        Month_Nm = "AUGUST"
    ElseIf Month(Date) = 9 Then
        Month_Nm = "SEPTEMBER"
    ElseIf Month(Date) = 10 Then
        Month_Nm = "OCTOBER"
    ElseIf Month(Date) = 11 Then
        Month_Nm = "NOVEMBER"
    ElseIf Month(Date) = 12 Then
        Month_Nm = "DECEMBER"
    End If
End Function

'====================================================='
'DATE COMPARE
'RETURN 1 WHEN FIRST DATE GREATER
'       2 WHEN SECOND GREATER
'       0(ZERO)WHEN BOTH EQUAL
Public Function DateCmp(ByVal FstDt As String, ByVal SecDt As String) As Integer
    If DatePart("yyyy", FstDt) > DatePart("yyyy", SecDt) Then
        DateCmp = 1
    ElseIf DatePart("yyyy", FstDt) < DatePart("yyyy", SecDt) Then
        DateCmp = 2
    ElseIf DatePart("m", FstDt) > DatePart("m", SecDt) Then
        DateCmp = 1
    ElseIf DatePart("m", FstDt) < DatePart("m", SecDt) Then
        DateCmp = 2
    ElseIf DatePart("d", FstDt) > DatePart("d", SecDt) Then
        DateCmp = 1
    ElseIf DatePart("d", FstDt) < DatePart("d", SecDt) Then
        DateCmp = 2
    Else
        DateCmp = 0
    End If
End Function

