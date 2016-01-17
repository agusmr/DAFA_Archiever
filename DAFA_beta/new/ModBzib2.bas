Attribute VB_Name = "ModBzib2"
' ***
' * libbz2.dll calling interface for VB
' *   coded by Arnout de Vries, Relevant Soft- & Mindware
' *   24 jan 2001
' *   Enjoy and use it as much as possible
' *
' * BZIP2 homepage: http://sourceware.cygnus.com/bzip2/
' * from the webpage:
' *    What is bzip2?
' *    bzip2 is a freely available, patent free (see below), high-quality data compressor.
' *    It typically compresses files to within 10% to 15% of the best available techniques
' *    (the PPM family of statistical compressors), whilst being around twice as fast at
' *    compression and six times faster at decompression.
' ***
Option Explicit

'Declares
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)
Private Declare Function z2Compress Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffCompress" (Dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal blockSize100k As Long, ByVal Verbosity As Long, ByVal workFactor As Long) As Long
Private Declare Function z2Decompress Lib "libbz2.dll" Alias "BZ2_bzBuffToBuffDecompress" (Dest As Any, destLen As Long, Source As Any, ByVal sourceLen As Long, ByVal Small As Long, ByVal Verbosity As Long) As Long

Public Function z2CompressData(TheData() As Byte, Hasil() As Byte, ByVal lCompressionLevel As Long) As Long

  'compressionlevel:
  ' 1 = superfast
  ' 9 = superthight
  
  'Allocate memory for byte array
  Dim BufferSize As Long
  Dim TempBuffer() As Byte
  Dim result As Long
  Dim lSourceLen As Long
  
  If lCompressionLevel > 9 Then lCompressionLevel = 9
  If lCompressionLevel < 0 Then lCompressionLevel = 1
  
  lSourceLen = UBound(TheData) + 1
  BufferSize = lSourceLen + (lSourceLen * 0.01) + 600
  ReDim TempBuffer(BufferSize)
  
  'Compress byte array (data)
  result = z2Compress(TempBuffer(0), BufferSize, TheData(0), lSourceLen, lCompressionLevel, 0, 0)
  
  'Truncate to compressed size
  ReDim Preserve Hasil(BufferSize - 1)
  CopyMemory Hasil(0), TempBuffer(0), BufferSize
  'Cleanup
  Erase TempBuffer
  
  'Set properties if no error occurred
  z2CompressData = UBound(Hasil) + 1
  
  'Return error code (if any)


End Function

Public Function z2DeCompressData(TheData() As Byte, Hasil() As Byte, lDestLen As Long) As Long

  'Allocate memory for byte array
  Dim TempBuffer() As Byte
  Dim result As Long
  Dim lSourceLen As Long
  Dim lVerbosity As Long ' We want the DLL to shut up, so set it to 0
  Dim lSmall As Long ' if <> 0 then use (s)low memory routines
  
  lVerbosity = 0
  lSmall = 0
  
  lSourceLen = UBound(TheData) + 1
  ReDim TempBuffer(lDestLen - 1)
  
  'Decompress byte array (data)
  result = z2Decompress(TempBuffer(0), lDestLen, TheData(0), lSourceLen, lSmall, lVerbosity)
  
  'Truncate to compressed size
  ReDim Preserve Hasil(lDestLen - 1)
  CopyMemory Hasil(0), TempBuffer(0), lDestLen
  
  'Cleanup
  Erase TempBuffer
  
  'Set properties if no error occurred
  z2DeCompressData = UBound(Hasil) + 1
  
  'Return error code (if any)
  

End Function


