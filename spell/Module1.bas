Attribute VB_Name = "Module1"
Private Function spell(ByVal num As Long) As String
Select Case num
            Case 1
                spell = " One"
            Case 2
                spell = " Two"
            Case 3
                spell = " Three"
            Case 4
                spell = " Four"
            Case 5
                spell = " Five"
            Case 6
                spell = " Six"
            Case 7
                spell = " Seven"
            Case 8
                spell = " Eight"
            Case 9
                spell = " Nine"
            Case 10
                spell = " Ten"
            Case 11
                spell = " Eleven"
            Case 12
                spell = " Twelve"
            Case 13
                spell = " Thirteen"
            Case 14
                spell = " Fourteen"
            Case 15
                spell = " Fifteen"
            Case 16
                spell = " Sixteen"
            Case 17
                spell = " Seventeen"
            Case 18
                spell = " Eighteen"
            Case 19
                spell = " Nineteen"
            Case 20
                spell = " Twenty"
            Case 30
                spell = " Thirty"
            Case 40
                spell = " Fourty"
            Case 50
                spell = " Fifty"
            Case 60
                spell = " Sixty"
            Case 70
                spell = " Seventy"
            Case 80
                spell = " Eighty"
            Case 90
                spell = " Ninety"
End Select
End Function
Public Function conwords(ByVal src_num As String) As String
Dim sno, nat As Double
sno = Val(src_num)
Dim words, whole As String
whole = sno
If sno < 1 Then words = " Zero"
nat = Val(whole)
If (Right(nat, 2)) > 0 And (Right(nat, 2)) < 21 Or (Right(nat, 2)) = 30 Or (Right(nat, 2)) = 40 Or (Right(nat, 2)) = 50 Or (Right(nat, 2)) = 60 Or (Right(nat, 2)) = 70 Or (Right(nat, 2)) = 80 Or (Right(nat, 2)) = 90 Then
    words = words + spell((Right(nat, 2)))
ElseIf (Right(nat, 2)) > 20 Then
    words = words + spell(Left(Right(nat, 2), 1) & "0")
    words = words + spell(Right(nat, 1))
End If
If nat > 99 Then
    If Left(Right(nat, 3), 1) <> 0 Then words = spell(Left(Right(nat, 3), 1)) + " Hundred" + words
End If
Count = Len(Trim(nat))
If nat > 999 Then
        If Mid(nat, Count - 4, 2) > 0 And Mid(nat, Count - 4, 2) < 21 Or Mid(nat, Count - 4, 2) = 30 Or Mid(nat, Count - 4, 2) = 40 Or Mid(nat, Count - 4, 2) = 50 Or Mid(nat, Count - 4, 2) = 60 Or Mid(nat, Count - 4, 2) = 70 Or Mid(nat, Count - 4, 2) = 80 Or Mid(nat, Count - 4, 2) = 90 Then
            words = spell(Mid(nat, Count - 4, 2)) + " Thousand" + words
        ElseIf Mid(nat, Count - 4, 2) > 20 Then
            words = spell(Mid(nat, Count - 4, 1) & "0") + spell(Right(Mid(nat, Count - 4, 2), 1)) + " Thousand" + words
        End If
End If
If nat > 99999 Then
        If Mid(nat, Count - 6, 2) > 0 And Mid(nat, Count - 6, 2) < 21 Or Mid(nat, Count - 6, 2) = 30 Or Mid(nat, Count - 6, 2) = 40 Or Mid(nat, Count - 6, 2) = 50 Or Mid(nat, Count - 6, 2) = 60 Or Mid(nat, Count - 6, 2) = 70 Or Mid(nat, Count - 6, 2) = 80 Or Mid(nat, Count - 6, 2) = 90 Then
            words = spell(Mid(nat, Count - 6, 2)) + " Lakh" + words
        ElseIf Mid(nat, Count - 6, 2) > 20 Then
            words = spell(Mid(nat, Count - 6, 1) & "0") + spell(Right(Mid(nat, Count - 6, 2), 1)) + " Lakh" + words
        End If
End If
If nat > 9999999 Then
        If Mid(nat, Count - 8, 2) > 0 And Mid(nat, Count - 8, 2) < 21 Or Mid(nat, Count - 8, 2) = 30 Or Mid(nat, Count - 8, 2) = 40 Or Mid(nat, Count - 8, 2) = 50 Or Mid(nat, Count - 8, 2) = 60 Or Mid(nat, Count - 8, 2) = 70 Or Mid(nat, Count - 8, 2) = 80 Or Mid(nat, Count - 8, 2) = 90 Then
            words = spell(Mid(nat, Count - 8, 2)) + " Crore" + words
        ElseIf Mid(nat, Count - 8, 2) > 20 Then
            words = spell(Mid(nat, Count - 8, 1) & "0") + spell(Right(Mid(nat, Count - 8, 2), 1)) + " Crore" + words
        End If
End If
conwords = words
End Function
