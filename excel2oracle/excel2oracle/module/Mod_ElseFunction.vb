Module Mod_ElseFunction

  Public Function Convert10to26(ByVal originalNum As Integer) As String
    If originalNum <= 0 Then Return ""
    Dim count As Integer = 0
    Dim strReturn As String = ""

    While originalNum > 0
      count = originalNum Mod 26
      If count = 0 Then
        count = 26
        originalNum = originalNum - 1
      End If
      strReturn = Convert.ToChar(64 + count) & strReturn
      originalNum = originalNum / 26

    End While
    Return strReturn
  End Function

End Module
