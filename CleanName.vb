Option Explicit On

Function CleanName(c As Range) As String

    Dim name As String
    Dim strPattern As String
    Dim regEx As New RegExp
    Dim match As Object

    name = c.Value                  ' get cell value
    strPattern = "[a-z]+,\s[a-z]+"  ' set pattern to match

    With regEx
        .Global = False             ' should only have 1 match
        .IgnoreCase = True
        .Pattern = strPattern
    End With

    match = regEx.Execute(name)     ' find match

    If match.Count <> 0 Then
        CleanName = match.Item(0)   ' return matched string
    Else
        CleanName = "None"
    End If

End Function