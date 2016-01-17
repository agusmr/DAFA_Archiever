Attribute VB_Name = "ModCrc16"
Dim CRCTab(255) As Long '               Array for single byte CRCs loaded from table
Const MaXWord As Long = 65535
Const MaXInt As Long = 32767

Private Sub CmdStart_Click()

End Sub
Public Function GetCrc16(DataByte() As Byte, ByVal Jum As Long) As Integer
    Dim x As Long
    Dim TempLng As Integer
    Dim CRC As Long
    
    CRC = &HFFFF
    
    For x = 0 To Jum
        TempLng = ((CRC \ 256) Xor DataByte(x)) '           Shift left (>>8) XOR with data
        CRC = ((CRC * 256) And 65535) Xor CRCTab(TempLng) ' Shift right (<<8) prevent overflow, XOR with table
    Next
    
    GetCrc16 = WordToInteger(CRC)

End Function

Public Sub InitCrc16()
    Dim x As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim CRC As Long

    For i = 0 To 255
      k = i * 256
      CRC = 0
      For j = 0 To 7
        If (((CRC Xor k) And 32768) = 32768) Then
          CRC = (CRC * 2) Xor &H1021
        Else
          CRC = (CRC * 2)
        End If
        k = k * 2
      Next
      CRCTab(i) = (CRC And 65535)
    Next i


End Sub

Public Function MyHash(z() As Byte, ByVal Jum As Long) As Integer
    Dim i As Long
    Dim temp As Long
    Dim CRC As Long
    Dim Tambah As Boolean
    
    CRC = 0
    Tambah = True
    For i = 0 To Jum
        temp = z(i)
        If Tambah = True Then
            If CRC <= &H7F00 Then
                CRC = CRC + temp
            Else
                CRC = CRC - temp
                Tambah = False
            End If
        Else
            If CRC >= &H80FF Then
                CRC = CRC - temp
            Else
                CRC = CRC + temp
                Tambah = True
            End If
        End If
    Next
    MyHash = CRC Xor &HFFFF
End Function

Private Function WordToInteger(ByVal Word As Long) As Integer
    Dim Tmp As Integer
    
    
    If Word > MaXInt Then
        Tmp = MaXWord - MaXInt - 1
    Else
        Tmp = Word
    End If
    WordToInteger = Tmp
End Function

