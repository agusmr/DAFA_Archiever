Attribute VB_Name = "ModAplib"
'---------------------------------------------------------------------------------------
' Name      : maPLib (Module)
'---------------------------------------------------------------------------------------
' Project   : aPLib Compression Library Visual Basic 6 Wrapper
' Author    : Jon Johnson
' Date      : 7/15/2005
' Email     : jjohnson@sherwoodpolice.org
' Version   : 1.0
' Purpose   : Wraps the functions in 'aplib.dll' for use in Visual Basic 6
' Notes     : The two functions (CompressFile, DecompressFile) have very limited error
'           : checking.  There is much room for improvement.
'---------------------------------------------------------------------------------------
Option Explicit

'---------------------------------------------------------------------------------------
' API Constants
'---------------------------------------------------------------------------------------
Const APLIB_ERROR = -1

'---------------------------------------------------------------------------------------
' API Compression Functions
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Declare   : aP_Pack
' Inputs    : source = Pointer to the data to be compressed.
'           : destination = Pointer to where the compressed data should be stored.
'           : length = The length of the uncompressed data in bytes.
'           : workmem = Pointer to the work memory which is used during compression.
'           : callback = Pointer to the callback function (or NULL).
'           : cbparam = Callback argument.
' Returns   : The length of the compressed data, or APLIB_ERROR on error.
' Purpose   : Compresses 'length' bytes of data from 'source()' into 'destination()',
'           : using 'workmem()' for temporary storage.
'---------------------------------------------------------------------------------------
Public Declare Function aP_Pack Lib "aplib.dll" Alias "_aP_pack" (source As Byte, destination As Byte, ByVal Length As Long, workmem As Byte, Optional ByVal callback As Long = &H0, Optional ByVal cbparam As Long = &H0) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_workmem_size
' Inputs    : input_size = The length of the uncompressed data in bytes.
' Returns   : The required length of the work buffer.
' Purpose   : Computes the required size of the 'workmem()' buffer used by 'aP_pack' for
'           : compressing 'input_size' bytes of data.
'---------------------------------------------------------------------------------------
Public Declare Function aP_workmem_size Lib "aplib.dll" Alias "_aP_workmem_size" (ByVal input_size As Long) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_max_packed_size
' Inputs    : input_size = The length of the uncompressed data in bytes.
' Returns   : The maximum possible size of the compressed data.
' Purpose   : Computes the maximum possible compressed size possible when compressing
'           : 'input_size' bytes of incompressible data.
'---------------------------------------------------------------------------------------
Public Declare Function aP_max_packed_size Lib "aplib.dll" Alias "_aP_max_packed_size" (ByVal input_size As Long) As Long
'---------------------------------------------------------------------------------------
' Declare   : aPsafe_pack
' Inputs    : source = Pointer to the data to be compressed.
'           : destination = Pointer to where the compressed data should be stored.
'           : length = The length of the uncompressed data in bytes.
'           : workmem = Pointer to the work memory which is used during compression.
'           : callback = Pointer to the callback function (or NULL).
'           : cbparam = Callback argument.
' Returns   : The length of the compressed data, or APLIB_ERROR on error.
' Purpose   : Wrapper function for 'aP_pack', which adds a header to the compressed data
'           : containing the length of the original data, and CRC32 checksums of the
'           : original and compressed data.
'---------------------------------------------------------------------------------------
Public Declare Function aPsafe_pack Lib "aplib.dll" Alias "_aPsafe_pack" (source As Byte, destination As Byte, ByVal Length As Long, workmem As Byte, Optional ByVal callback As Long = &H0, Optional ByVal cbparam As Long = &H0) As Long


'---------------------------------------------------------------------------------------
' API Decompression Functions
'---------------------------------------------------------------------------------------
'---------------------------------------------------------------------------------------
' Declare   : aP_depack
' Inputs    : source = Pointer to the compressed data.
'           : destination = Pointer to where the decompressed data should be stored.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Decompresses the compressed data from 'source()' into 'destination()'.
'---------------------------------------------------------------------------------------
Public Declare Function aP_depack Lib "aplib.dll" Alias "_aP_depack_asm_fast" (source As Byte, destination As Byte) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_depack_safe
' Inputs    : source = Pointer to the compressed data.
'           : srclen = The size of the source buffer in bytes.
'           : destination = Pointer to where the decompressed data should be stored.
'           : dstlen = The size of the destination buffer in bytes.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Decompresses the compressed data from 'source()' into 'destination()'.
'---------------------------------------------------------------------------------------
Public Declare Function aP_depack_safe Lib "aplib.dll" Alias "_aP_depack_asm_safe" (source As Byte, ByVal srcLen As Long, destination As Byte, ByVal dstlen As Long) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_depack_asm
' Inputs    : source = Pointer to the compressed data.
'           : destination = Pointer to where the decompressed data should be stored.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Decompresses the compressed data from 'source()' into 'destination()'.
'---------------------------------------------------------------------------------------
Public Declare Function aP_depack_asm Lib "aplib.dll" Alias "_aP_depack_asm" (source As Byte, destination As Byte) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_depack_asm_fast
' Inputs    : source = Pointer to the compressed data.
'           : destination = Pointer to where the decompressed data should be stored.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Decompresses the compressed data from 'source()' into 'destination()'.
'---------------------------------------------------------------------------------------
Public Declare Function aP_depack_asm_fast Lib "aplib.dll" Alias "_aP_depack_asm_fast" (source As Byte, destination As Byte) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_depack_asm_safe
' Inputs    : source = Pointer to the compressed data.
'           : srclen = The size of the source buffer in bytes.
'           : destination = Pointer to where the decompressed data should be stored.
'           : dstlen = The size of the destination buffer in bytes.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Decompresses the compressed data from 'source()' into 'destination()'.
'---------------------------------------------------------------------------------------
Public Declare Function aP_depack_asm_safe Lib "aplib.dll" Alias "_aP_depack_asm_safe" (source As Byte, ByVal srcLen As Long, destination As Byte, ByVal dstlen As Long) As Long
'---------------------------------------------------------------------------------------
' Declare   : aP_crc32
' Inputs    : source = Pointer to the data to process.
'           : length = The size in bytes of the data.
' Returns   : The CRC32 value.
' Purpose   : Computes the CRC32 value of 'length' bytes of data from 'source()'.
'---------------------------------------------------------------------------------------
Public Declare Function aP_crc32 Lib "aplib.dll" Alias "_aP_crc32" (source As Byte, ByVal Length As Long) As Long
'---------------------------------------------------------------------------------------
' Declare   : aPsafe_check
' Inputs    : source = The compressed data to process.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Computes the CRC32 of the compressed data in 'source()' and checks it
'           : against the value in the header.  Returns the length of the decompressed
'           : data stored in the header.
'---------------------------------------------------------------------------------------
Public Declare Function aPsafe_check Lib "aplib.dll" Alias "_aPsafe_check" (source As Byte) As Long
'---------------------------------------------------------------------------------------
' Declare   : aPsafe_get_orig_size
' Inputs    : source = The compressed data to process.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Returns the length of the decompressed data stored in the header of the
'           : compressed data in 'source()'.
'---------------------------------------------------------------------------------------
Public Declare Function aPsafe_get_orig_size Lib "aplib.dll" Alias "_aPsafe_get_orig_size" (source As Byte) As Long
'---------------------------------------------------------------------------------------
' Declare   : aPsafe_depack
' Inputs    : source = Pointer to the compressed data.
'           : srclen = The size of the source buffer in bytes.
'           : destination = Pointer to where the decompressed data should be stored.
'           : dstlen = The size of the destination buffer in bytes.
' Returns   : The length of the decompressed data, or APLIB_ERROR on error.
' Purpose   : Wrapper function for 'aP_depack_asm_safe', which checks the CRC32 of the
'           : compressed data, decompresses, and checks the CRC32 of the decompressed
'           : data.
'---------------------------------------------------------------------------------------
Public Declare Function aPsafe_depack Lib "aplib.dll" Alias "_aPsafe_depack" (source As Byte, ByVal srcLen As Long, destination As Byte, ByVal dstlen As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

'---------------------------------------------------------------------------------------
' Procedure : CompressFile
' Returns   : Boolean (True if succesful, False if not)
' DateTime  : 7/15/2005
' Author    : Jon Johnson <jjohnson@sherwoodpolice.org>
' Purpose   : Example of using aPLib to compress a file
'---------------------------------------------------------------------------------------
Public Function CompressByte1(bIn() As Byte, bOut() As Byte) As Long
    Dim lSize As Long     'Length of compressed data
    Dim bWork() As Byte       'Work buffer
    

    ReDim bWork(aP_workmem_size(UBound(bIn) + 1))
    
    ReDim bOut(aP_max_packed_size(UBound(bIn) + 1))
    
    lSize = aPsafe_pack(bIn(0), bOut(0), (UBound(bIn) + 1), bWork(0))
    'lSize = aPsafe_pack2(VarPtr(bIn(0)), VarPtr(bOut(0)), (UBound(bIn) + 1), VarPtr(bWork(0)))
    'lSize = Invoke("aplib.dll", "_aPsafe_pack", bIn(0), bOut(0), (UBound(bIn) + 1), bWork(0), 0, 0)
    If lSize = APLIB_ERROR Then
        CompressByte1 = False
        Exit Function
    End If
    
    ReDim Preserve bOut(lSize - 1)

    'Everything went OK, return True
    CompressByte1 = lSize
End Function
Public Function DecompressByte1(bIn() As Byte, bOut() As Byte, ByVal szUk As Long) As Long
    Dim lSize As Long     'Length of compressed data
        
    ReDim bOut(szUk - 1)
    
    lSize = aPsafe_depack(bIn(0), (UBound(bIn) + 1), bOut(0), szUk)
    
    If lSize = APLIB_ERROR Then
        DecompressByte1 = False
        Exit Function
    End If

    'Everything went OK, return True
    DecompressByte1 = lSize
End Function
