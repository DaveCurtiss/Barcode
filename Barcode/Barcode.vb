Public Class Code128

    Public Function Encode128(Input As String) As String
        Dim i As Integer
        Dim output As String = ""

        'Check for zero length string
        If Len(Input) = 0 Then
            Return ""
        End If

        'Check for valid ASCII input
        For i = 1 To Input.Length
            Select Case AscW(Mid(Input, i, 1))
                Case 32 To 126
                Case Else
                    i = 0
                    Exit For
            End Select
        Next

        If i <> 0 Then

            output = Strings.Left(Input, InStr(Input, " "))
            If OddDigits(Input) Then
                'use Table B for 1st digit if necessary and then shift to Table C
                output &= Mid(Input, InStr(Input, " ") + 1, 1) & ChrW(199)
            End If

            'convert digits to ASCII representation of Code128 character
            For i = InStr(Input, " ") + 2 To Len(Input) Step 2
                If Mid(Input, i, 2) <= 94 Then
                    output &= ChrW(Mid(Input, i, 2) + 32)
                Else
                    output &= ChrW(Mid(Input, i, 2) + 105)
                End If
            Next

            'add startcode, generate check character and stop code
            output &= CalculateCheckChar(output) & ChrW(206)
            output = ChrW(204) & output
        End If

        Return output

    End Function

    Private Function OddDigits(Input As String) As Boolean
        Dim i As Integer
        Dim digits As Integer = 0
        Dim result As Boolean = False
        For i = 1 To Input.Length
            If IsNumeric(Mid(Input, i, 1)) Then
                digits += 1
            End If
        Next
        If digits Mod 2 <> 0 Then
            result = True
        End If
        Return result

    End Function

    Private Function CalculateCheckChar(Input As String) As String
        Dim i As Integer
        Dim sumOfChars As Integer = 104
        Dim Result As Integer

        For i = 1 To Input.Length
            If AscW(Mid(Input, i, 1)) < 127 Then
                sumOfChars += (AscW(Mid(Input, i, 1)) - 32) * i
            Else
                sumOfChars += (AscW(Mid(Input, i, 1)) - 100) * i
            End If
        Next

        Result = sumOfChars Mod 103
        If Result <= 94 Then
            Return ChrW(Result + 32)
        Else
            Return ChrW(Result + 100)
        End If


    End Function

End Class
