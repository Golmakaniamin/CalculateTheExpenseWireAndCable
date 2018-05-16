Attribute VB_Name = "KontrolGardeshmes"
Dim infogardesh(29, 1) As String
Dim infogardesh1(29, 1) As String
Dim countsql As String

Public Sub amin1(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      With Form9.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 +1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»«‰ç—"
              infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
              
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
        Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
        
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin2(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form10.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 +1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»«‰ç—"
              infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
              
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
        Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
        
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin3(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form11.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 +1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»«‰ç—"
              infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
              
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
        Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
        
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin4(q As Integer)
  For q = 0 To 29
    infogardesh(q, 0) = 0
    infogardesh(q, 1) = 0
  Next q
      
      With Form13.Adodc3
        .Recordset.MoveFirst
        Do
          Select Case .Recordset.Fields!Name
            Case "À«‰ÊÌÂ"
              infogardesh(0, 0) = Val(infogardesh(0, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(0, 1) = Val(infogardesh(0, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "‰Â«ÌÌ"
              infogardesh(1, 0) = Val(infogardesh(1, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(1, 1) = Val(infogardesh(1, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "òÊ—Â"
              infogardesh(2, 0) = Val(infogardesh(2, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(2, 1) = Val(infogardesh(2, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case " «»"
              infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 6 +1"
              infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 36 + 1"
              infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«” —‰œ— 4 + 1"
              infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "œ—«„  ÊÌ” —"
              infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "„Œ«»—« Ì"
              infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«ò” —Êœ—"
              infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»” Â »‰œÌ"
              infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "«‰»«— „Õ’Ê·"
              infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
            Case "»«‰ç—"
              infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(.Recordset.Fields!naghlbebadmeghdar)
              infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(.Recordset.Fields!naghlbebadmoney)
              
          End Select
          .Recordset.MoveNext
        Loop Until .Recordset.EOF = True
        
        Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
        Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
        Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
        Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
        Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
        Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
        Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
        Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
        Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
        Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
        Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
        Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
        Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
        Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
        Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
        Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
        Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
        Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
        Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
        Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
        Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
        
        For q = 0 To 28
          infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
          infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
        Next q
        Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
        Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
        Form23.Adodc1.Recordset.Update
      End With
End Sub

Public Sub amin5(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = " «»"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Taab WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin6(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "«” —‰œ— 6 +1"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Sterander1_6 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin7(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "«” —‰œ— 36 + 1"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Sterander1_36 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin8(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "«” —‰œ— 4 + 1"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Sterander1_4 WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin9(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "œ—«„  ÊÌ” —"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From DramToester WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin10(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "„Œ«»—« Ì"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Mokhaberat WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin11(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "«ò” —Êœ—"
db1.Open Form3.Text10.Text

For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
For q = 0 To 29
  infogardesh1(q, 0) = 0
  infogardesh1(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Exteroder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(3, 0) = Val(infogardesh1(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(3, 1) = Val(infogardesh1(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
             
            Case "«” —‰œ— 6 +1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(4, 0) = Val(infogardesh1(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(4, 1) = Val(infogardesh1(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 36 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(5, 0) = Val(infogardesh1(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(5, 1) = Val(infogardesh1(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 4 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(6, 0) = Val(infogardesh1(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(6, 1) = Val(infogardesh1(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
            
            Case "œ—«„  ÊÌ” —"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(7, 0) = Val(infogardesh1(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(7, 1) = Val(infogardesh1(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "„Œ«»—« Ì"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(8, 0) = Val(infogardesh1(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(8, 1) = Val(infogardesh1(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
                        
            Case "«ò” —Êœ—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(9, 0) = Val(infogardesh1(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(9, 1) = Val(infogardesh1(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»” Â »‰œÌ"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(10, 0) = Val(infogardesh1(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(10, 1) = Val(infogardesh1(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«‰»«— „Õ’Ê·"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(11, 0) = Val(infogardesh1(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(11, 1) = Val(infogardesh1(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»«‰ç—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(12, 0) = Val(infogardesh1(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(12, 1) = Val(infogardesh1(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  
  For q = 0 To 28
    infogardesh1(29, 0) = Val(infogardesh1(29, 0)) + Val(infogardesh1(q, 0))
    infogardesh1(29, 1) = Val(infogardesh1(29, 1)) + Val(infogardesh1(q, 1))
  Next q
  
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin12(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "»” Â »‰œÌ"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
For q = 0 To 29
  infogardesh1(q, 0) = 0
  infogardesh1(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Bastebandi WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(3, 0) = Val(infogardesh1(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(3, 1) = Val(infogardesh1(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
             
            Case "«” —‰œ— 6 +1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(4, 0) = Val(infogardesh1(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(4, 1) = Val(infogardesh1(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 36 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(5, 0) = Val(infogardesh1(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(5, 1) = Val(infogardesh1(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 4 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(6, 0) = Val(infogardesh1(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(6, 1) = Val(infogardesh1(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
            
            Case "œ—«„  ÊÌ” —"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(7, 0) = Val(infogardesh1(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(7, 1) = Val(infogardesh1(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "„Œ«»—« Ì"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(8, 0) = Val(infogardesh1(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(8, 1) = Val(infogardesh1(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
                        
            Case "«ò” —Êœ—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(9, 0) = Val(infogardesh1(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(9, 1) = Val(infogardesh1(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»” Â »‰œÌ"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(10, 0) = Val(infogardesh1(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(10, 1) = Val(infogardesh1(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«‰»«— „Õ’Ê·"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(11, 0) = Val(infogardesh1(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(11, 1) = Val(infogardesh1(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»«‰ç—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(12, 0) = Val(infogardesh1(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(12, 1) = Val(infogardesh1(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  
  For q = 0 To 28
    infogardesh1(29, 0) = Val(infogardesh1(29, 0)) + Val(infogardesh1(q, 0))
    infogardesh1(29, 1) = Val(infogardesh1(29, 1)) + Val(infogardesh1(q, 1))
  Next q
  
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin13(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "«‰»«— „Õ’Ê·"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From AnbarMahsol WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin14(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "»«‰ç—"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Bancher WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
             
            Case "«” —‰œ— 6 +1"
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 36 + 1"
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«” —‰œ— 4 + 1"
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
            
            Case "œ—«„  ÊÌ” —"
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "„Œ«»—« Ì"
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
                        
            Case "«ò” —Êœ—"
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»” Â »‰œÌ"
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "«‰»«— „Õ’Ê·"
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
            Case "»«‰ç—"
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin15(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "«ò” —Êœ—"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
  infogardesh1(q, 0) = 0
  infogardesh1(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Exteroder WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(3, 0) = Val(infogardesh1(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(3, 1) = Val(infogardesh1(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
             
            Case "«” —‰œ— 6 +1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(4, 0) = Val(infogardesh1(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(4, 1) = Val(infogardesh1(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 36 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(5, 0) = Val(infogardesh1(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(5, 1) = Val(infogardesh1(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 4 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(6, 0) = Val(infogardesh1(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(6, 1) = Val(infogardesh1(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
            
            Case "œ—«„  ÊÌ” —"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(7, 0) = Val(infogardesh1(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(7, 1) = Val(infogardesh1(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "„Œ«»—« Ì"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(8, 0) = Val(infogardesh1(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(8, 1) = Val(infogardesh1(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
                        
            Case "«ò” —Êœ—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(9, 0) = Val(infogardesh1(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(9, 1) = Val(infogardesh1(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»” Â »‰œÌ"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(10, 0) = Val(infogardesh1(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(10, 1) = Val(infogardesh1(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«‰»«— „Õ’Ê·"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(11, 0) = Val(infogardesh1(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(11, 1) = Val(infogardesh1(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»«‰ç—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(12, 0) = Val(infogardesh1(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(12, 1) = Val(infogardesh1(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh1(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh1(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh1(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh1(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh1(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh1(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh1(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh1(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh1(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh1(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh1(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh1(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh1(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh1(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh1(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh1(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh1(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh1(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh1(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh1(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh1(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh1(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh1(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh1(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh1(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh1(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  
  For q = 0 To 28
    infogardesh1(29, 0) = Val(infogardesh1(29, 0)) + Val(infogardesh1(q, 0))
    infogardesh1(29, 1) = Val(infogardesh1(29, 1)) + Val(infogardesh1(q, 1))
  Next q
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh1(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh1(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

Public Sub amin16(q As Integer)
Dim db1 As New ADODB.Connection
Dim rs1(2) As New ADODB.Recordset
Dim strstep As String

strstep = "»” Â »‰œÌ"
db1.Open Form3.Text10.Text
For q = 0 To 29
  infogardesh(q, 0) = 0
  infogardesh(q, 1) = 0
Next q
  
For q = 0 To 29
  infogardesh1(q, 0) = 0
  infogardesh1(q, 1) = 0
Next q
  
  rs1(1).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (name='" + strstep + "')", db1
  countsql = rs1(1).Fields!number1
  rs1(1).Close
  
  rs1(1).Open "SELECT * FROM ozanmasir WHERE (name='" + strstep + "')", db1
  If countsql > 0 Then
    rs1(1).MoveFirst
    Do
      rs1(0).Open "SELECT count(rad) as number1 FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ", db1
      countsql = rs1(0).Fields!number1
      rs1(0).Close
      
      rs1(0).Open "SELECT * FROM ozanmasir WHERE (idmahsol=" + Trim(Str(rs1(1).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(1).Fields!rad)) + ") ORDER BY rad1 ASC", db1
      If countsql > 0 Then
        rs1(0).Find "name= '" + strstep + "'", , adSearchForward, 1
        rs1(0).MoveNext
        If rs1(0).EOF = False Then
          rs1(2).Open "SELECT * From Bastebandi WHERE (idmahsol=" + Trim(Str(rs1(0).Fields!idmahsol)) + ") and (rad=" + Trim(Str(rs1(0).Fields!rad)) + ")", db1
          Select Case rs1(0).Fields!Name
            Case " «»"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(3, 0) = Val(infogardesh(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(3, 1) = Val(infogardesh(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(3, 0) = Val(infogardesh1(3, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(3, 1) = Val(infogardesh1(3, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
             
            Case "«” —‰œ— 6 +1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(4, 0) = Val(infogardesh(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(4, 1) = Val(infogardesh(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(4, 0) = Val(infogardesh1(4, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(4, 1) = Val(infogardesh1(4, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 36 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(5, 0) = Val(infogardesh(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(5, 1) = Val(infogardesh(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(5, 0) = Val(infogardesh1(5, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(5, 1) = Val(infogardesh1(5, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«” —‰œ— 4 + 1"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(6, 0) = Val(infogardesh(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(6, 1) = Val(infogardesh(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(6, 0) = Val(infogardesh1(6, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(6, 1) = Val(infogardesh1(6, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
            
            Case "œ—«„  ÊÌ” —"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(7, 0) = Val(infogardesh(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(7, 1) = Val(infogardesh(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(7, 0) = Val(infogardesh1(7, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(7, 1) = Val(infogardesh1(7, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "„Œ«»—« Ì"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(8, 0) = Val(infogardesh(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(8, 1) = Val(infogardesh(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(8, 0) = Val(infogardesh1(8, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(8, 1) = Val(infogardesh1(8, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
                        
            Case "«ò” —Êœ—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(9, 0) = Val(infogardesh(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(9, 1) = Val(infogardesh(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(9, 0) = Val(infogardesh1(9, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(9, 1) = Val(infogardesh1(9, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»” Â »‰œÌ"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(10, 0) = Val(infogardesh(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(10, 1) = Val(infogardesh(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(10, 0) = Val(infogardesh1(10, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(10, 1) = Val(infogardesh1(10, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "«‰»«— „Õ’Ê·"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(11, 0) = Val(infogardesh(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(11, 1) = Val(infogardesh(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(11, 0) = Val(infogardesh1(11, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(11, 1) = Val(infogardesh1(11, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
            Case "»«‰ç—"
              If rs1(2).Fields!nomes = "„”  Ê·Ìœ ‘œÂ" Then
                infogardesh(12, 0) = Val(infogardesh(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh(12, 1) = Val(infogardesh(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              Else
                infogardesh1(12, 0) = Val(infogardesh1(12, 0)) + Val(rs1(2).Fields!naghlbebadmeghdar)
                infogardesh1(12, 1) = Val(infogardesh1(12, 1)) + Val(rs1(2).Fields!naghlbebadmoney)
              End If
              
          End Select
          rs1(2).Close
        End If
      End If
      rs1(0).Close
      rs1(1).MoveNext
    Loop Until rs1(1).EOF = True
  End If
  rs1(1).Close
  
  Form23.Adodc1.Recordset.Fields!sanaveye1 = infogardesh1(0, 0)
  Form23.Adodc1.Recordset.Fields!sanaveye2 = infogardesh1(0, 1)
  Form23.Adodc1.Recordset.Fields!nahaee1 = infogardesh1(1, 0)
  Form23.Adodc1.Recordset.Fields!nahaee2 = infogardesh1(1, 1)
  Form23.Adodc1.Recordset.Fields!Koreh1 = infogardesh1(2, 0)
  Form23.Adodc1.Recordset.Fields!Koreh2 = infogardesh1(2, 1)
  Form23.Adodc1.Recordset.Fields!Taab1 = infogardesh1(3, 0)
  Form23.Adodc1.Recordset.Fields!Taab2 = infogardesh1(3, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_61 = infogardesh1(4, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_62 = infogardesh1(4, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_361 = infogardesh1(5, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_362 = infogardesh1(5, 1)
  Form23.Adodc1.Recordset.Fields!Sterander1_41 = infogardesh1(6, 0)
  Form23.Adodc1.Recordset.Fields!Sterander1_42 = infogardesh1(6, 1)
  Form23.Adodc1.Recordset.Fields!DramToester1 = infogardesh1(7, 0)
  Form23.Adodc1.Recordset.Fields!DramToester2 = infogardesh1(7, 1)
  Form23.Adodc1.Recordset.Fields!Mokhaberat1 = infogardesh1(8, 0)
  Form23.Adodc1.Recordset.Fields!Mokhaberat2 = infogardesh1(8, 1)
  Form23.Adodc1.Recordset.Fields!Exteroder1 = infogardesh1(9, 0)
  Form23.Adodc1.Recordset.Fields!Exteroder2 = infogardesh1(9, 1)
  Form23.Adodc1.Recordset.Fields!Bastebandi1 = infogardesh1(10, 0)
  Form23.Adodc1.Recordset.Fields!Bastebandi2 = infogardesh1(10, 1)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol1 = infogardesh1(11, 0)
  Form23.Adodc1.Recordset.Fields!AnbarMahsol2 = infogardesh1(11, 1)
  Form23.Adodc1.Recordset.Fields!bancher1 = infogardesh1(12, 0)
  Form23.Adodc1.Recordset.Fields!bancher2 = infogardesh1(12, 1)
  
  For q = 0 To 28
    infogardesh(29, 0) = Val(infogardesh(29, 0)) + Val(infogardesh(q, 0))
    infogardesh(29, 1) = Val(infogardesh(29, 1)) + Val(infogardesh(q, 1))
  Next q
  
  For q = 0 To 28
    infogardesh1(29, 0) = Val(infogardesh1(29, 0)) + Val(infogardesh1(q, 0))
    infogardesh1(29, 1) = Val(infogardesh1(29, 1)) + Val(infogardesh1(q, 1))
  Next q
  
  Form23.Adodc1.Recordset.Fields!sum1 = Val(infogardesh1(29, 0))
  Form23.Adodc1.Recordset.Fields!sum2 = Val(infogardesh1(29, 1))
  Form23.Adodc1.Recordset.Update
db1.Close
End Sub

