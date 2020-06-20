Public Class Code39

    Public Function Encode39(input As String) As String
        Dim output As String = ""

        'Check for zero length string
        If Len(input) = 0 Then
            Return ""
        End If

        'Check for valid ASCII input
        If ValidASCII(input) = False Then
            Return ""
        End If

        output = "*" & input.ToUpper & "*"
        Return output

    End Function
    Private Function ValidASCII(input As String) As Boolean
        Dim i As Integer
        Dim result As Boolean = False

        For i = 1 To input.Length
            Select Case AscW(Mid(input, i, 1))
                Case 32 To 126 'ensure input is valid ASCII
                    result = True
                Case Else
                    result = False
                    Exit For
            End Select
        Next
        Return result

    End Function

End Class
