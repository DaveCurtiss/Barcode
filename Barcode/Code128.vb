Public Class Code128

    Public Function Encode128(Input As String) As String
        Dim i As Integer
        Dim output As String = ""
        Dim startCode As String = ""
        Dim UseTableB As Boolean = True

        'Check for zero length string
        If Len(Input) = 0 Then
            Return ""
        End If

        'Check for valid ASCII input
        If ValidASCII(Input) = False Then
            Return ""
        End If

        i = 1
        While i <= Len(Input)
            If i = 1 Then 'check if 1st 4 characters are digits and set startcode
                If LookAheadForDigits(Input, i, 4) = True Then
                    startCode = ChrW(205) 'start in table C
                    UseTableB = False
                Else
                    startCode = ChrW(204) 'start in table B
                    UseTableB = True
                End If
            Else
                If i + 3 = Len(Input) Then 'check if last 4 characters are digits
                    If LookAheadForDigits(Input, i, 4) = True Then
                        output &= ChrW(199) 'switch to Table C
                        UseTableB = False
                    Else
                        UseTableB = True
                    End If
                Else
                    If LookAheadForDigits(Input, i, 6) = True Then 'check for 6 consecutive digits
                        output &= ChrW(199) 'switch to Table C
                        UseTableB = False
                    Else
                        UseTableB = True
                    End If
                End If
            End If
            If UseTableB = False Then 'using table C, try to process 2 digits
                If LookAheadForDigits(Input, i, 2) = True Then
                    If Mid(Input, i, 2) <= 94 Then
                        output &= ChrW(Mid(Input, i, 2) + 32)
                    Else
                        output &= ChrW(Mid(Input, i, 2) + 105)
                    End If
                    i += 2
                Else
                    output &= ChrW(200) 'switch to Table B
                    i += 1
                End If
            Else 'using table B, process one character
                output &= Mid(Input, i, 1)
                i += 1
            End If
        End While

        'add startcode, generate check character and stopcode
        output &= CalculateCheckChar(output) & ChrW(206)
        output = startCode & output

        Return output

    End Function

    Private Function LookAheadForDigits(input As String, start As Integer, length As Integer) As Boolean
        'returns true if the number of characters to look ahead for are numeric
        Dim i As Integer
        Dim result As Boolean = False

        For i = 1 To length
            If start + i <= Len(input) AndAlso IsNumeric(Mid(input, start, i)) Then
                result = True
            Else
                result = False
                Exit For
            End If
        Next
        Return result

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
