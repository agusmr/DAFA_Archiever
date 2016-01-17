Attribute VB_Name = "ModLZ"

'Author         : Agus Minanur Rohman
'Filename       : Checksum Adler32.bas (cCRC32 Class Module)
'Description    : Calculate Adler32 Checksum of a string
'Date           : Rabu, 15 Desember, 2010, 15:00

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private AsmCode() As Byte
Const Adler_Code As String = ""
' Recoded Adler32 source Asm By ManiaX Code
Private Declare Function blz_pack Lib "brieflz.dll" Alias "_blz_pack@16" (source As Byte, _
                 destination As Byte, _
                 ByVal size_t As Long, _
                 workmem As Byte) As Long
Private Declare Function blz_max_packed Lib "brieflz.dll" Alias "_blz_max_packed_size@4" (ByVal size_t As Long) As Long
Private Declare Function blz_workmem_size Lib "brieflz.dll" Alias "_blz_workmem_size@4" (ByVal input_size As Long) As Long
Private Declare Function blz_depack Lib "brieflz.dll" Alias "_blz_depack@12" (source As Byte, destination As Byte, ByVal depacked_size As Long) As Long
Public Function CompressByte2(BIn() As Byte, bOut() As Byte) As Long
    Dim lSize As Long     'Length of compressed data
    Dim bWork() As Byte       'Work buffer
    

    ReDim bWork(blz_workmem_size(UBound(BIn) + 1))
    
    ReDim bOut(blz_max_packed(UBound(BIn) + 1))
    
    lSize = blz_pack(BIn(0), bOut(0), (UBound(BIn) + 1), bWork(0))
    
    If lSize = APLIB_ERROR Then
        CompressByte2 = False
        Exit Function
    End If
    
    ReDim Preserve bOut(lSize - 1)

    'Everything went OK, return True
    CompressByte2 = lSize
End Function
Public Function DecompressByte2(BIn() As Byte, bOut() As Byte, ByVal szWal As Long) As Long
    Dim lSize As Long     'Length of compressed data
    Dim szOut As Long
    
    
    ReDim bOut(szWal - 1)
    
    lSize = blz_depack(BIn(0), bOut(0), szWal)
    
    If lSize = APLIB_ERROR Then
        DecompressByte2 = False
        Exit Function
    End If

    'Everything went OK, return True
    DecompressByte2 = szWal
End Function

Public Function Press(BIn() As Byte, bOut() As Byte) As Long
    Dim SizeSource As Long
    Dim source As Long
    Dim Code As Long
    Dim lSize As Long     'Length of compressed data
    Dim bWork() As Byte       'Work buffer
    

    
    'If Len(Text) > 0 Then
        Code = VarPtr(AsmCode(0))
        source = VarPtr(BIn(0))
        
        SizeSource = UBound(BIn) + 1
        
        ReDim bWork(blz_workmem_size(SizeSource))
        ReDim bOut(blz_max_packed(SizeSource))
        
        Press = CallWindowProc(Code, source, VarPtr(bOut(0)), SizeSource, VarPtr(bWork(0)))
        ReDim Preserve bOut(Press - 1)
    'End If
    
End Function


Public Sub InitPack()
    Dim x1$, x2$, x3$, x4$, x5$, x6$, Temp$
    'Initialize Pack precompiled assembly code
    
    x1 = "5356575583EC08FC8B7C242831C0B900000400F3AB8B74241C8B7C24208B5C24248D4433FC8944240489342485DB0F84820200008A064688074783FB010F847302000066BD010089FA83C702EB1A85ED750689FA4583C7026601ED730566892A31ED8A06468807473B7424040F83110200008B4C242889F38B342429F3510F"
    x2 = "B60669C03D0100000FB64E0101C869C03D0100000FB64E0201C869C03D0100000FB64E0301C825FFFF030059893481464B75CB893424510FB60669C03D0100000FB64E0101C869C03D0100000FB64E0201C869C03D0100000FB64E0301C825FFFF0300598B1C8185DB0F8460FFFFFF8B4C240429F183C1045231C08A14033A"
    x3 = "14067504404975F45A83F8040F823EFFFFFF89F129D985ED750689FA4583C7026601ED6645730566892A31ED01C683E8025350D1E8BB01000000D1E8744811DBEBF8721685ED750689FA4583C7026601ED730566892A31EDEB1685ED750689FA4583C7026601ED6645730566892A31ED85ED750689FA4583C7026601ED6645"
    x4 = "730566892A31EDD1EB75B858D1E8721685ED750689FA4583C7026601ED730566892A31EDEB1685ED750689FA4583C7026601ED6645730566892A31ED85ED750689FA4583C7026601ED730566892A31ED5B4989C8C1E80883C0025350D1E8BB01000000D1E8744811DBEBF8721685ED750689FA4583C7026601ED730566892A"
    x5 = "31EDEB1685ED750689FA4583C7026601ED6645730566892A31ED85ED750689FA4583C7026601ED6645730566892A31EDD1EB75B858D1E8721685ED750689FA4583C7026601ED730566892A31EDEB1685ED750689FA4583C7026601ED6645730566892A31ED85ED750689FA4583C7026601ED730566892A31ED5B880F473B74"
    x6 = "24040F82EFFDFFFF8B5C240483C304EB1A85ED750689FA4583C7026601ED730566892A31ED8A064688074739DE72E285ED74086601ED73FB66892A89F82B44242083C4085D5F5E5BC21000"
     
    Temp = x1 & x2 & x3 & x4 & x5 & x6
    ReDim AsmCode(Len(Temp) \ 2 - 1)
    
    For i = 1 To Len(Temp) Step 2
        AsmCode(i \ 2) = Val("&H" & Mid$(Temp, i, 2))
    Next i

End Sub

'Created by Agus Minanur R.| A.k.a| ManiaX Code| Copyright © 2010 DarmaSoft, indonesia


