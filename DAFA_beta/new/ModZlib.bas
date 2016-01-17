Attribute VB_Name = "ModZlib"
'Enum CZErrors 'for compression/decompression
'    Z_OK = 0
'    Z_STREAM_END = 1
'    Z_NEED_DICT = 2
'    Z_ERRNO = -1
'    Z_STREAM_ERROR = -2
'    Z_DATA_ERROR = -3
'    Z_MEM_ERROR = -4
'    Z_BUF_ERROR = -5
'    Z_VERSION_ERROR = -6
'End Enum

'Enum CompressionLevels 'for compression/decompression
'    Z_NO_COMPRESSION = 0
'    Z_BEST_SPEED = 1
'    'note that levels 2-8 exist, too
'    Z_BEST_COMPRESSION = 9
'    Z_DEFAULT_COMPRESSION = -1
'End Enum

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function Compress Lib "Zlib.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
Private Declare Function compress2 Lib "Zlib.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long, ByVal Level As Long) As Long
Private Declare Function uncompress Lib "Zlib.dll" (Dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Long
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

