Function SymbolTxtConverter(str, Optional NoDotReplace As Integer = 0)
 
    If (NoDotReplace = 1) Then
        strCharToReplace = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ-/,;:!@#$%¨&*()[]{}´`'_=+|\º<>" & Chr(13) & Chr(10)
        strCharReplacement = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN--------------------------------"
    Else
        strCharToReplace = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ-/,;:!@#$%¨&*()[]{}´`'_=+|\.º<>" & Chr(13) & Chr(10)
        strCharReplacement = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN---------------------------------"
    End If
    
    tmp = str
    
    For i = 1 To Len(tmp)
        p = InStr(strCharToReplace, Mid(tmp, i, 1))
        If p > 0 Then Mid(tmp, i, 1) = Mid(strCharReplacement, p, 1)
    Next
    
    SymbolTxtConverter = tmp
     
End Function
