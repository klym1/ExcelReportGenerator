Imports System.IO

Class Myh

    Public Function Page_Load(ByVal sender As Object, ByVal fName As String) As String
        Dim fStream As FileStream
        Dim bReader As BinaryReader
        Dim uBuffer As UInt16
        Dim result As String
        Dim bFindBOF As Boolean = True
        Dim bBuffer As Byte
        Dim j As Integer = 0
        If (String.IsNullOrEmpty(fName)) Then
            Return String.Empty
        End If
        'Read XLS file
        Try
            fStream = New FileStream(fName, FileMode.Open, FileAccess.Read)
            bReader = New BinaryReader(fStream)
            Do While j < fStream.Length - 1
                If bFindBOF = True Then
                    'Get BOF Record
                    Do While j < fStream.Length - 3 And bFindBOF = True
                        uBuffer = bReader.ReadUInt16()
                        If uBuffer = 2057 Then 'BOF = 0x0809
                            bFindBOF = False
                            result += "ID = 0x0809 | "
                            'Response.Write("ID = 0x0809 | ")
                        End If
                        j = j + 2
                    Loop
                    If bFindBOF = False Then
                        'Get Record Length
                        uBuffer = bReader.ReadUInt16()
                        result += ("Length = " & uBuffer.ToString() & _
                        " | Offset = " & j.ToString())
                        j = j + 2
                        'Get Record Body
                        result += (IIf(uBuffer > 0, " | Body = ", String.Empty))
                        For i = 1 To uBuffer
                            bBuffer = bReader.ReadByte()
                            result += (GetByteHex(bBuffer) & " | ")
                        Next
                        result += ("<br/>")
                        j = j + uBuffer
                    Else
                        j = j + 2
                    End If
                Else
                    'Get Record ID
                    uBuffer = bReader.ReadUInt16()
                    If uBuffer = 10 Then 'EOF = 0x000A
                        bFindBOF = True
                    End If
                    result += ("ID = " & GetUInt16Hex(uBuffer) & " | ")
                    j = j + 2
                    'Get Record Length
                    uBuffer = bReader.ReadUInt16()
                    result += ("Length = " & uBuffer.ToString() & _
                    " | Offset = " & j.ToString())
                    j = j + 2
                    'Get Record Body
                    result += (IIf(uBuffer > 0, " | Body = ", String.Empty))
                    For i = 1 To uBuffer
                        bBuffer = bReader.ReadByte()
                        result += (GetByteHex(bBuffer) & " | ")
                    Next
                    result += ("<br/>")
                    j = j + uBuffer
                End If
            Loop
            result += ("*ID and LENGTH take 4 bytes, " & _
            "so total bytes per record is 4 + Length.")
            bReader.Dispose()
            fStream.Close()
        Catch ex As Exception
            result += ("xls-check::Page_Load: exception: " & _
            ex.Message & Environment.NewLine & ex.StackTrace)
        End Try
    End Function

    Private Function GetUInt16Hex(ByVal pValue As UInt16) As String
        Dim sResult As String = "0x"
        Dim dTemp As Double = 0
        '16^3
        dTemp = Math.Floor(pValue / 4096)
        sResult = sResult & GetHexLetter(dTemp).ToString()
        pValue = pValue - (dTemp * 4096)
        '16^2
        dTemp = Math.Floor(pValue / 256)
        sResult = sResult & GetHexLetter(dTemp).ToString()
        pValue = pValue - (dTemp * 256)
        '16^1
        dTemp = Math.Floor(pValue / 16)
        sResult = sResult & GetHexLetter(dTemp).ToString()
        pValue = pValue - (dTemp * 16)
        '16^0
        sResult = sResult & GetHexLetter(pValue).ToString()
        Return sResult
    End Function

    Private Function GetByteHex(ByVal pValue As UInt16) As String
        Dim sResult As String = "0x"
        Dim dTemp As Double = 0
        '16^1
        dTemp = Math.Floor(pValue / 16)
        sResult = sResult & GetHexLetter(dTemp).ToString()
        pValue = pValue - (dTemp * 16)
        '16^0
        sResult = sResult & GetHexLetter(pValue).ToString()
        Return sResult
    End Function

    Private Function GetHexLetter(ByVal pValue As Double) As String
        Dim sResult As String = String.Empty
        If pValue < 10 Then
            sResult = pValue.ToString()
        ElseIf pValue = 10 Then
            sResult = "A"
        ElseIf pValue = 11 Then
            sResult = "B"
        ElseIf pValue = 12 Then
            sResult = "C"
        ElseIf pValue = 13 Then
            sResult = "D"
        ElseIf pValue = 14 Then
            sResult = "E"
        ElseIf pValue = 15 Then
            sResult = "F"
        End If
        Return sResult
    End Function
End Class
