Attribute VB_Name = "i_pM_CRC32"
Option Explicit
Option Private Module

Public Function sCRC32Hash(ByVal sInput As String) As String
    Dim laCRC32Table(255) As Long
    Dim i As Long, j As Long
    Dim lCRC As Long
    Dim bytaByteArray() As Byte

   On Error GoTo Catch
    ' Initialize CRC32 table
    For i = 0 To 255
        lCRC = i
        For j = 0 To 7
            If (lCRC And 1) Then
                lCRC = &HEDB88320 Xor (lCRC \ 2)
            Else
                lCRC = lCRC \ 2
            End If
        Next j
        laCRC32Table(i) = lCRC
    Next i

    ' Convert input to byte array
    bytaByteArray = StrConv(sInput, vbFromUnicode)

    ' Calculate CRC32 value
    lCRC = &HFFFFFFFF
    For i = LBound(bytaByteArray) To UBound(bytaByteArray)
        lCRC = laCRC32Table((lCRC Xor bytaByteArray(i)) And &HFF) Xor (lCRC \ 256)
    Next i
    lCRC = Not lCRC

    ' Konvertiere den CRC32-Wert in eine hexadezimale Zeichenkette
    sCRC32Hash = LCase(Format$(lCRC, "00000000"))
    Exit Function
Catch:
   sCRC32Hash = "Error"
End Function

