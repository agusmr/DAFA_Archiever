Attribute VB_Name = "ModRC4"


' Module Encription RC4
'-------------------------------------------------------------------------------------
'Author         : Ron Rivest RSA
'Coded          : Agus Minanur Rohman a.k.a ManiaX Code
'Filename       : RC4.cls (RC4 Class Module)
'Description    : Module Encription Algorithm RC4
'Date           : 19 Agustus 2009
'-------------------------------------------------------------------------------------
 
'-------------------------------------------------------------------------------------
' Penting [ Dilarang keras menghilangkan Nama Penulis ]

Dim Bit() As Byte
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (destination As Any, source As Any, ByVal Length As Long)
Private s(0 To 255) As Integer
Private U(0 To 255) As Integer
Public Sub DecriptBit(ByRef Bit() As Byte, Optional kunci As String)
'Lakukan Encript berdasarkan kunci
    Call EncriptBit(Bit(), kunci)
End Sub
Public Function Encript(ByRef Bit() As Byte, Optional kunci As String) As Long
    Call EncriptBit(Bit(), kunci)
End Function
Public Function Decript(ByRef Bit() As Byte, Optional kunci As String) As Long
    Call DecriptBit(Bit(), kunci)
End Function
Sub Tukar(ByRef a As Integer, ByRef b As Integer)
'Tukar A menjadi B dan B menjadi A
    Dim t As Integer
    t = a
    a = b
    b = t
End Sub
Public Sub EncriptBit(ByRef Bit() As Byte, Optional kunci As String)
Dim j As Long, b As Long, t As Byte, KEY() As Byte, i As Long

    KEY() = StrConv(kunci, vbFromUnicode)
    
'Proses inisialisasi S-Box (Array S)
    For i = 0 To 255
        s(i) = i
    Next i

'Proses inisialisasi S-Box (Array U)
    For i = 0 To 255
        U(i) = KEY(i Mod Len(kunci))
    Next i
    
'Kemudian melakukan langkah pengacakan S-Box
    For i = 0 To 255
    
        b = (b + s(i) + U(i)) Mod 256
        Tukar s(i), s(b)
    Next i
 
'Hanya untuk memindah Byte di Memorry panjang 512 Byte dan memasukkan ke S()
Call CopyMem(s(0), s(0), 512)

'Setelah itu membuat pseudo random byte
    For y = 0 To (UBound(Bit))
        i = (i + 1) Mod 256
        j = (j + s(i)) Mod 256
        
        Tukar s(i), s(j)
    
        t = (s((s(i) + s(j)) Mod 256))
        k = Bit(y)
        Bit(y) = k Xor t
    Next y
End Sub



