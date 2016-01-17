Attribute VB_Name = "ModInvoke"
'---------------------------------------------------------------------------------------
' Module    : mZombieInvoke
' Author    : Karcrack
' Now       : 09/08/2010 13:37
' Purpose   : Calling API without declaring
'             Only uses VB6 functions :)
' History   : 20100908 First cut .......................................................
'---------------------------------------------------------------------------------------

Option Explicit

Private Type Zombie_STRUCT1
    cNull       As Currency 'Must be 0
    ppS2        As Long 'Pointer to pointer to Zombie_STRUCT2
End Type

Private Type Zombie_STRUCT2
    lNull       As Long 'Must be 0
    lAddr       As Long 'The Addr
End Type

Private Type tAPICall
    ptsLIB      As Long ' Pointer to ANSI String that contains Library (NULL TERMINATED!)
    ptsProc     As Long ' Pointer to ANSI String that contains Procedure(NULL TERMINATED!)
    lReserved   As Long ' Just reserved...
    lPointer    As Long ' Pointer to the buffer that will contain temp variables from DllFunctionCall
    lpBuffer(3) As Long ' Buffer that will contain temp variables
End Type

Private Type DUMB_LONG
    lLNG        As Long
End Type

Private Type BYTES_LONG
    b1          As Byte:    b2          As Byte
    b3          As Byte:    b4          As Byte
End Type

'MSVBVM60
Private Declare Function DllFunctionCall Lib "MSVBVM60" (ByRef typeAPI As tAPICall) As Long
Private Declare Function Zombie_AddRef Lib "MSVBVM60" (ByRef tStructure As Zombie_STRUCT1) As Long

Private bvASM(&HFF) As Byte

Public Function Invoke(ByVal sLibName As String, ByVal sProcName As String, ParamArray vParams() As Variant) As Long
    Dim hMod        As Long
    Dim S1          As Zombie_STRUCT1
    Dim S2          As Zombie_STRUCT2
    Dim i           As Long
    Dim iCount      As Long

    hMod = GetPointer(sLibName, sProcName)

    '//POP EAX                  '//POP EBX                  '//PUSH EAX
    Call AddByte(&H58, iCount): Call AddByte(&H5B, iCount): Call AddByte(&H50, iCount)

    For i = UBound(vParams) To LBound(vParams) Step -1
        '//PUSH CLng(vParams(i))
        Call AddPush(CLng(vParams(i)), iCount)
    Next i

    '//CALL hMod                '//RET
    Call AddCall(hMod, iCount): Call AddByte(&HC3, iCount)

    S2.lAddr = VarPtr(bvASM(0))
    S1.ppS2 = VarPtr(VarPtr(S2))

    Invoke = Zombie_AddRef(S1)
End Function

Private Function GetPointer(ByVal sLib As String, ByVal sProc As String) As Long
    Dim tAPI        As tAPICall
    Dim bvLib()     As Byte
    Dim bvMod()     As Byte

    bvLib = StrConv(sLib + vbNullChar, vbFromUnicode):  bvMod = StrConv(sProc + vbNullChar, vbFromUnicode)

    With tAPI
        .ptsLIB = VarPtr(bvLib(0)):     .ptsProc = VarPtr(bvMod(0))
        .lReserved = &H40000:           .lPointer = VarPtr(.lpBuffer(0))
    End With

    GetPointer = DllFunctionCall(tAPI)
End Function

Private Sub AddCall(ByVal lpPtrCall As Long, ByRef iCount As Long)
    Call AddByte(&HB8, iCount)                  '//MOV EAX, ________
    Call AddLong(lpPtrCall, iCount)             '//_______, XXXXXXXX
    Call AddByte(&HFF, iCount)                  '//CALL EXX
    Call AddByte(&HD0, iCount)                  '//____ EAX
End Sub

Private Sub AddPush(ByVal lLong As Long, ByRef iCount As Long)
    Call AddByte(&H68, iCount)                  '//PUSH, ________
    Call AddLong(lLong, iCount)                 '//____, XXXXXXXX
End Sub

Private Sub AddLong(ByVal lLong As Long, ByRef iCount As Long)
    'Swap Endian (Ej: 0xDEADBEEF <-> 0xEFBEADDE)
    Dim tDL         As DUMB_LONG
    Dim tBL         As BYTES_LONG

    tDL.lLNG = lLong
    LSet tBL = tDL

    Call AddByte(tBL.b1, iCount):   Call AddByte(tBL.b2, iCount)
    Call AddByte(tBL.b3, iCount):   Call AddByte(tBL.b4, iCount)
End Sub

Private Sub AddByte(ByVal bByte As Byte, ByRef iCount As Long)
    bvASM(iCount) = bByte:    iCount = iCount + 1
End Sub

