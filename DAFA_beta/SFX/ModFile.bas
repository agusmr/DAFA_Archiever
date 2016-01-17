Attribute VB_Name = "ModFile"
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function CopyFile Lib "kernel32" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As Any) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As Any) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Public Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Private zFileName   As String
Private hFile       As Long
Private nFileLen    As Long
Private nOperation  As Long
Public Function SetFileAttr(ByVal Alamat As String, ByVal Tipe As Integer) As Boolean
    SetFileAttr = SetFileAttributes(StrPtr(Alamat), Tipe)
End Function
Public Function VbReadFileB(ByVal hFile As Long, ByVal Awal As Long, ByVal Panjang As Long, ByRef isi() As Byte) As Long
    Erase isi
    Dim ret        As Long
    
    DoEvents
    ReDim isi(Panjang - 1) As Byte
    SetFilePointer hFile, (Awal - 1), 0, 0
    Call ReadFile(hFile, isi(0), Panjang, ret, ByVal 0&)
End Function
Public Function HapusFile(ByVal Alamat As String) As Long
    SetFileAttributes StrPtr(Alamat), &H80
    HapusFile = DeleteFile(StrPtr(Alamat))
End Function
Public Function VbCloseHandle(ByVal Handle As Long) As Long
    VbCloseHandle = CloseHandle(Handle)
End Function
Public Function VbOpenFile(ByVal szFileName As String) As Long
    VbOpenFile = CreateFileW(StrPtr(szFileName), &H80000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
End Function
Public Function VbFileLen(ByVal nFileHandle As Long) As Long
    VbFileLen = GetFileSize(nFileHandle, 0)
End Function
Public Function SetAttrib(ByVal Alamat As String) As Long
    Call SetFileAttributes(StrPtr(Alamat), &H80)
End Function
Public Function VbWriteFileB(ByVal Handle As Long, ByVal Panjang As Long, ByRef Bit() As Byte) As Long
    Dim ret As Long
    
    If Panjang > 0 Then
        Call WriteFile(Handle, Bit(0), Panjang, ret, ByVal 0&)
    End If
End Function

