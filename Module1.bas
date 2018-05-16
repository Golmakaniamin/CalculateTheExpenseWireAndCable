Attribute VB_Name = "Aminmath"
Public Function aminint(ByVal a As String) As String
Dim q As String
If a = 0 Then aminint = 0: Exit Function
If InStr(a, ".") = 0 Then aminint = a: Exit Function

If Mid(a, InStr(a, ".") + 1, 1) >= 5 Then
  q = Mid(a, 1, Val(InStr(a, ".") - 1))
  aminint = Val(q) + 1
Else
  aminint = Mid(a, 1, Val(InStr(a, ".") - 1))
End If
End Function
