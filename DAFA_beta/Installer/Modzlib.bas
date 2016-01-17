Attribute VB_Name = "Modzlib"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Compress Lib "compres.dll" Alias "compress" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "compres.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal Level As Long) As Long
Private Declare Function uncompress Lib "compres.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Dim lngCompressedSize As Long
Dim lngDecompressedSize As Long

Public Function CompressByte(Bit() As Byte, Out() As Byte, Level As Long) As Long
    Dim lngResult As Long
    Dim JumByte As Long
    Dim ByteArray() As Byte
    Dim UkWal As Long
    
    UkWal = UBound(Bit) + 1
    JumByte = UkWal * 1.01 + 12
    
    ReDim ByteArray(JumByte)
    lngResult = compress2(ByteArray(0), JumByte, Bit(0), UkWal, Level)
    
    ReDim Out(JumByte - 1)
    CopyMemory Out(0), ByteArray(0), JumByte
    Erase ByteArray
    
    CompressByte = JumByte
End Function
Public Function DecompressByteArray(TheData() As Byte, Out() As Byte, OriginalSize As Long) As Long
    Dim lngResult As Long
    Dim lngBufferSize As Long
    Dim arrByteArray() As Byte
    
    lngDecompressedSize = OriginalSize
    lngCompressedSize = UBound(TheData) + 1
    lngBufferSize = OriginalSize
    lngBufferSize = lngBufferSize * 1.01 + 12
    
    ReDim arrByteArray(lngBufferSize)
    lngResult = uncompress(arrByteArray(0), lngBufferSize, TheData(0), lngCompressedSize)
    
    ReDim Preserve Out(lngBufferSize - 1)
    CopyMemory Out(0), arrByteArray(0), lngBufferSize
    Erase arrByteArray
    DecompressByteArray = lngBufferSize
End Function


