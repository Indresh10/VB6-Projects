Attribute VB_Name = "Book"
Public conn As New ADODB.Connection
Public bkType As String, userType As String, userNm As String
Public srchCategory As String

Public Sub Main()
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Library.mdb;Persist Security Info=False"
    FrmWelcome.Show
End Sub

'LOCK TEXTBOXES OF BOOK ENTRY FORM
Public Sub lockText(Frm As Form)
    Frm.TxtCode.Locked = True
    Frm.TxtTitle.Locked = True
    Frm.TxtAuther.Locked = True
    Frm.TxtPub.Locked = True
    Frm.CmbDay.Locked = True
    Frm.CmbMonth.Locked = True
    Frm.CmbYear.Locked = True
    Frm.TxtPrice.Locked = True
    Frm.TxtQty.Locked = True
    Frm.TxtFrom.Locked = True
End Sub

'DISABLE COMMAND BUTTON
Public Sub disableCommand(Frm As Form)
    Frm.CmdFirst.Enabled = False
    Frm.CmdPrv.Enabled = False
    Frm.CmdNext.Enabled = False
    Frm.CmdLast.Enabled = False
    
    Frm.CmdAdd.Enabled = False
    Frm.CmdEdit.Enabled = False
    Frm.CmdDel.Enabled = False
    Frm.CmdSave.Enabled = False
    Frm.FremCategory.Enabled = False
End Sub
'ENABLE COMMAND BUTTON
Public Sub enableCommand(Frm As Form)
    Frm.CmdFirst.Enabled = True
    Frm.CmdPrv.Enabled = True
    Frm.CmdNext.Enabled = True
    Frm.CmdLast.Enabled = True
    
    Frm.CmdAdd.Enabled = True
    Frm.CmdEdit.Enabled = True
    Frm.CmdDel.Enabled = True
    Frm.CmdSave.Enabled = True
    Frm.FremCategory.Enabled = True
End Sub

'ENABLE TEXT BOXES IN BOOK ENTRY FORM
Public Sub enableText(Frm As Form)
    'frm.TxtCode.Locked = False
    Frm.TxtTitle.Locked = False
    Frm.TxtAuther.Locked = False
    Frm.TxtPub.Locked = False
    Frm.CmbDay.Locked = False
    Frm.CmbMonth.Locked = False
    Frm.CmbYear.Locked = False
    Frm.TxtPrice.Locked = False
    Frm.TxtQty.Locked = False
    Frm.TxtFrom.Locked = False
End Sub

'CLEAR TEXT BOXES OF BOOK ENTRY
Public Sub clearText(Frm As Form)
    Frm.TxtCode.Text = ""
    Frm.TxtTitle.Text = ""
    Frm.TxtAuther.Text = ""
    Frm.TxtPub.Text = ""
    Frm.CmbDay.Text = Day(Date)
    Frm.CmbMonth.Text = Month(Date)
    Frm.CmbYear.Text = Year(Date)
    Frm.TxtPrice.Text = 0
    Frm.TxtQty.Text = 0
    Frm.TxtFrom.Text = ""
    Frm.TxtAvlQty.Text = 0
End Sub

'RETRIVE RECORD IN BOOK ENTRY FORM
Public Sub bookData(Frm As Form, ByVal rs As Recordset)
        Frm.CmdSave.Enabled = False
        'RETRIVE RECORD
        Frm.TxtCode.Text = rs.Fields(0)
        Frm.TxtTitle.Text = rs.Fields(1)
        Frm.TxtAuther.Text = rs.Fields(2)
        Frm.TxtPub.Text = rs.Fields(3)
        Frm.CmbDay.Text = Day(rs.Fields(4))
        Frm.CmbMonth.Text = Month(rs.Fields(4))
        Frm.CmbYear.Text = Year(rs.Fields(4))
        Frm.TxtPrice.Text = rs.Fields(5)
        Frm.TxtQty.Text = rs.Fields(6)
        Frm.TxtAvlQty.Text = rs.Fields(6) - rs.Fields(8)
        Frm.TxtFrom.Text = rs.Fields(7)
End Sub

'TO UPPER
Public Function upper(Key As Integer)
    upper = Asc(UCase(Chr(Key)))
End Function

'SELECT TEXT
Public Sub selectTxt(Txt As TextBox)
    Txt.SelStart = 0
    Txt.SelLength = Len(Txt.Text)
End Sub

'BOOK CODE VALIDATION
Public Function codeValid(str As String, title As String, typ As String) As Boolean
    If Len(str) < 6 Then
        MsgBox "Invalid Code number. Code must be six character long", vbCritical, title
        codeValid = True
    Else
        codeValid = False
    End If
End Function

'DATA VALIDATION
Public Function dataValid(Frm As Form, title As String) As Boolean
    Dim result As Boolean
    result = False
    
    'WHEN BOOK CODE IS NOT STARTED WITH 'B' OR 'C'
    If Mid(Frm.TxtCode, 1, 1) <> "B" And Mid(Frm.TxtCode, 1, 1) <> "C" Then
        MsgBox "Invalid Code. Code must start with 'B' if book or 'C' if CD", vbInformation, title
        result = True
        dataValid = True
        Exit Function
    End If
    
    'VALIDATION IF ALL REQUIRED DATA ARE ENTERED OR NOT
    If Trim(Frm.TxtTitle) = "" Then
        result = True
    ElseIf Trim(Frm.TxtPub) = "" Then
        result = True
    ElseIf Trim(Frm.TxtPrice) = 0 Then
        result = True
    ElseIf Trim(Frm.TxtQty) = 0 Then
        result = True
    End If
    
    If result = True Then
        MsgBox "Please fill compulsory fields.", vbInformation, title
        dataValid = True
    Else
        dataValid = False
    End If
End Function

'====================================================
'PROCEDURE FOR CREATING NEXT CODE FOR BOOK/CD
Public Function Next_Code(ByRef rs As Recordset, typ As String) As String
    Dim tmp As String ', typ As String
    
    If typ = "BOOK" Then
        typ = "B"
    ElseIf typ = "CD" Then
        typ = "C"
    Else
        typ = "M"
    End If
    'GENERATING NEXT CODE FOR BOOK/CD
        If rs.RecordCount > 0 Then
            rs.MoveLast
            tmp = Val(Mid(rs.Fields(0), 2, 5)) + 1
        Else
            tmp = 1
        End If
        
        Select Case Len(tmp)
            Case 1
                Next_Code = typ & "0000" & tmp
            Case 2
                Next_Code = typ & "000" & tmp
            Case 3
                Next_Code = typ & "00" & tmp
            Case 4
                Next_Code = typ & "0" & tmp
            Case 5
                Next_Code = typ & tmp
        End Select
    
End Function

'=======================================================
'CODE TO CREATE LABEL OF BOOK SEARCH FORM
Public Sub searchLabel(Frm As Form)
    Dim lbl As String
    With Frm
        'SELECT TYPE
        If .OptBook.Value Then
            lbl = .OptBook.Caption
        Else
            lbl = .OptCd.Caption
        End If
        'SELECT SEARCH CATEGORY
        If .OptCode.Value Then
            lbl = lbl & " " & .OptCode.Caption
        ElseIf (.OptAuther.Value) Then
            lbl = lbl & " " & .OptAuther.Caption
        ElseIf (.OptTitle.Value) Then
            lbl = lbl & " " & .OptTitle.Caption
        ElseIf (.OptPublisher.Value) Then
            lbl = lbl & " " & .OptPublisher.Caption
        End If
        .LblSearch.Caption = lbl & " : "
    End With
End Sub

'===========================================================
'FILL BOOK/CD SEARCH FlexGrid
Public Sub fillGrid(Frm As Form, opt As String, ByVal rs As Recordset)
    Dim r As Integer, c As Integer
    Dim rs_tmp As New ADODB.Recordset
    
    With Frm
        If rs_tmp.State = 1 Then
            rs_tmp.Close
        End If
    
        If bkType = "BOOK" Then
            rs_tmp.Open "select * from Book_Mast where code like('B%') order by " & opt, conn, adOpenStatic, adLockPessimistic
        Else
            rs_tmp.Open "select * from Book_Mast where code like('C%') order by " & opt, conn, adOpenStatic, adLockPessimistic
        End If
        
        'SET TITLE OF MSFlexGrid
        .MsfgSearch.Cols = rs_tmp.Fields.Count
        .MsfgSearch.Rows = rs_tmp.RecordCount + 1

        'FILL MSFlexGrid
        If rs_tmp.RecordCount = 0 Then
            Exit Sub
        End If
        
        rs_tmp.MoveFirst
        For r = 1 To rs_tmp.RecordCount
            .MsfgSearch.TextMatrix(r, 0) = r
            .MsfgSearch.TextMatrix(r, 1) = rs_tmp.Fields(0)
            .MsfgSearch.TextMatrix(r, 2) = rs_tmp.Fields(1)
            .MsfgSearch.TextMatrix(r, 3) = rs_tmp.Fields(2)
            .MsfgSearch.TextMatrix(r, 4) = rs_tmp.Fields(3)
            .MsfgSearch.TextMatrix(r, 5) = Format(rs_tmp.Fields(4), "dd-mm-yyyy")
            .MsfgSearch.TextMatrix(r, 6) = rs_tmp.Fields(5)
            .MsfgSearch.TextMatrix(r, 7) = rs_tmp.Fields(6)
            .MsfgSearch.TextMatrix(r, 8) = rs_tmp.Fields(6) - rs_tmp.Fields(8)
            rs_tmp.MoveNext
        Next
        
    End With
End Sub

'===================================================
'TO FILL BOOK GRID WHEN SEARCHING
Public Sub fillGrid1(Frm As Form, rs As Recordset)
    Dim r As Integer
    Frm.MsfgSearch.Cols = rs.Fields.Count
    Frm.MsfgSearch.Rows = rs.RecordCount + 1
    
    If rs.RecordCount > 0 Then
    rs.MoveFirst
    For r = 1 To rs.RecordCount
        Frm.MsfgSearch.TextMatrix(r, 0) = r
        Frm.MsfgSearch.TextMatrix(r, 1) = rs.Fields(0)
        Frm.MsfgSearch.TextMatrix(r, 2) = rs.Fields(1)
        Frm.MsfgSearch.TextMatrix(r, 3) = rs.Fields(2)
        Frm.MsfgSearch.TextMatrix(r, 4) = rs.Fields(3)
        Frm.MsfgSearch.TextMatrix(r, 5) = rs.Fields(4)
        Frm.MsfgSearch.TextMatrix(r, 6) = rs.Fields(5)
        Frm.MsfgSearch.TextMatrix(r, 7) = rs.Fields(6)
        Frm.MsfgSearch.TextMatrix(r, 8) = rs.Fields(6) - rs.Fields(8)
        rs.MoveNext
    Next
    End If
End Sub


