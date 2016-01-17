Attribute VB_Name = "ModTemp"
Private Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Const MAX_PATH = 260
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Public Property Get TempDir() As String
    Dim sRet As String, c As Long
    Dim lErr As Long
    
   sRet = String$(MAX_PATH, 0)
   c = GetTempPath(MAX_PATH, sRet)
   lErr = Err.LastDllError
   If c = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   TempDir = Left$(sRet, c) & "DAFA\"
End Property
Public Function TempFile(Optional ByVal sPrefix As String, Optional ByVal sPathName As String) As String
    Dim lErr As Long
    Dim iPos As Long
    
   If sPrefix = "" Then sPrefix = ""
   If sPathName = "" Then sPathName = TempDir
   Dim sRet As String
   sRet = String(MAX_PATH, 0)
   GetTempFileName sPathName, sPrefix, 0, sRet
   lErr = Err.LastDllError
   If Not lErr = 0 Then
      Err.Raise 10000 Or lErr, App.EXEName & ".cAniCursor", WinAPIError(lErr)
   End If
   iPos = InStr(sRet, vbNullChar)
   If Not iPos = 0 Then
      TempFile = Left$(sRet, iPos - 1)
   End If
End Function
Private Function WinAPIError(ByVal lLastDLLError As Long) As String
    Dim sBuff As String
    Dim lCount As Long
    
   sBuff = String$(256, 0)
   lCount = FormatMessage( _
      FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
      0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
   If lCount Then
      WinAPIError = Left$(sBuff, lCount)
   End If
End Function




