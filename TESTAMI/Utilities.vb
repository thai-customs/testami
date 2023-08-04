Public Class Utilities
    '----------------------------------------------------------------
    ' Str2Bin() : Convert HEX STR to BINARY
    '----------------------------------------------------------------

    Public Function Str2Bin(ByVal StrIn As String, ByRef BinOut() As Byte, ByRef BinLen As Integer) As String
        Dim StrLen, pos, i As Short
        Dim StrTmp As String
        Dim ArrSize As Short
        Dim Int1, Int3, Int2 As Short

        StrLen = Len(StrIn)
        pos = 1

        ArrSize = StrLen / 2

        ReDim BinOut(ArrSize)

        For i = 0 To ArrSize - 1
            StrTmp = Mid(StrIn, pos, 1)
            HexToDec(StrTmp, Int1)
            pos = pos + 1

            StrTmp = Mid(StrIn, pos, 1)
            HexToDec(StrTmp, Int2)
            pos = pos + 1

            Int3 = (Int1 * 16) + Int2
            BinOut(i) = CByte(Int3)
            'Debug.Print BinOut(i)
        Next
        BinLen = ArrSize

        Return ""
    End Function

    Public Function Bin2Str(ByRef BinIn As String, ByVal BinInSize As Integer) As String
        Dim i As Short
        Dim tstr1, StrOut As String

        StrOut = Space(BinInSize * 2)

        For i = 0 To BinInSize - 1
            tstr1 = Right("0" & Hex(Asc(Mid(BinIn, i + 1, 1))), 2)
            StrOut = Trim(StrOut) & tstr1
        Next i

        Bin2Str = StrOut

    End Function

    '----------------------------------------------------------------
    ' HexToDec() : Convert HEX to Integer
    '----------------------------------------------------------------

    Public Function HexToDec(ByRef iHex As String, ByRef oDec As Short) As String

        Select Case iHex
            Case "A", "a"
                oDec = 10
            Case "B", "b"
                oDec = 11
            Case "C", "c"
                oDec = 12
            Case "D", "d"
                oDec = 13
            Case "E", "e"
                oDec = 14
            Case "F", "f"
                oDec = 15
            Case Else
                oDec = Val(iHex)
        End Select

        Return ""
    End Function
End Class
