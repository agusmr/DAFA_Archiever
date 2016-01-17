Attribute VB_Name = "ModKumpul"
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As MSG) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As MSG) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Public Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
Private Declare Function FindFirstFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindNextFileW Lib "kernel32" (ByVal hFindFile As Long, ByVal lpFindFileData As Long) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Const INVALID_HANDLE_VALUE = -1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const PM_REMOVE = &H1
Private Const vbKeyDot = 46
Private Const MaxLen = 260

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type POINTAPI
   x As Long
   y As Long
End Type

Private Type MSG
   hwnd     As Long
   Message  As Long
   wParam   As Long
   lParam   As Long
   Time     As Long
   pt       As POINTAPI
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MaxLen
    cShortFileName As String * 14
End Type

Private Wfd As WIN32_FIND_DATA
Private FileSpec As String
Private UseFileSpec As Boolean
Private hFindFile As Long
Private Message As MSG
Private hFile As Long
Private Const DOT1 As String = "."
Private Const DOT2 As String = ".."
Private Const Bintang As String = "*"
Private Const Slas As String = "\"
Private Const Semua As String = "*.*"
Dim G As cListItem

Public Function CallSearch(ByRef dirs As Long, ByRef dirbuff() As String, ByRef i As Long, UseIt As Boolean)
    Dim Temp As String
    If hFindFile <> INVALID_HANDLE_VALUE Then
        Do
            If Skip = True Or StopScan = True Then Exit Function
            If Wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                If AscW(Wfd.cFileName) <> vbKeyDot Then
                    If (dirs And 9) = 0 Then ReDim Preserve dirbuff(dirs + 10)
                    dirs = dirs + 1
                    JumDir = JumDir + 1
                    dirbuff(dirs) = StripNulls(Wfd.cFileName)
                End If
            ElseIf UseIt Then
                JumFile = JumFile + 1
            End If
            myDoEvents
        Loop While FindNextFileW(hFindFile, VarPtr(Wfd))
        Call FindClose(hFindFile)
    End If
End Function
Public Sub SearchFile(ByVal PathSearch As String)
    Dim dirs As Long, dirbuff() As String, i As Long
    
        hFindFile = FindFirstFileW(StrPtr(PathSearch & Semua), VarPtr(Wfd))
        CallSearch dirs, dirbuff(), i, True
        
        For i = 1 To dirs
            Call SearchFile(PathSearch & dirbuff(i) & Slas)
        Next
        
        If (JumFile And &HFF) = &HFF Then
            With frmProses
                .lblFile.Caption = "Jumlah File  : " & CStr(JumFile)
                .lblFolder.Caption = "Jumlah Folder : " & CStr(JumDir)
            End With
        End If
End Sub
Public Sub Kumpulkan(ByVal Path As String, ByVal hFileW As Long, Optional partPath As String = "")
    Dim dirs As Long, dirbuff() As String, i As Long
    
        hFindFile = FindFirstFileW(StrPtr(Path & Bintang), VarPtr(Wfd))
        Call CallSearch(dirs, dirbuff(), i, False)
        hFindFile = FindFirstFileW(StrPtr(Path & Semua), VarPtr(Wfd))
        If hFindFile <> INVALID_HANDLE_VALUE Then
            Do
                Call Olah(1, Path, Wfd, hFileW, partPath)
                
            Loop While FindNextFileW(hFindFile, VarPtr(Wfd))
            
            Call FindClose(hFindFile)
        End If
        
        For i = 1 To dirs
            Call Kumpulkan(Path & dirbuff(i) & Slas, hFileW, partPath)
        Next
Ex:
Exit Sub
    Call FindClose(hFindFile)
End Sub
Public Function Olah(ByVal Pilihan As Long, ByVal Path As String, Info As WIN32_FIND_DATA, hFileW As Long, Optional partPath As String = "") As Long
    Dim FileName As String
    Dim Attr As Long
    Dim FullPath As String
    
    FileName = StripNulls(Info.cFileName)
    FullPath = Path & FileName
    ukuRan = Info.nFileSizeLow
    Attr = Info.dwFileAttributes
        
        If PathIsDirectory(StrPtr(FullPath)) Then
            If Right$(FullPath, 3) = "\.." Then
                FullPath = AmbilAlamat2(FullPath)
            Else
                GoTo loncat
            End If
        End If
        Location = Replace(FullPath, partPath, "")
        
        If hFileW <> -1 Then
            TulisFile hFileW, Location, FullPath, 0, Attr
        Else
            Set G = FrmUtama.LVSelect.ListItems.Add(, FullPath)
                G.SubItem(2).Text = 1
                G.SubItem(3).Text = Location
        End If
loncat:
    
    myDoEvents
End Function
Public Function myDoEvents() As Boolean
      If PeekMessage(Message, 0, 0, 0, PM_REMOVE) Then
         TranslateMessage Message
         DispatchMessage Message
         myDoEvents = True
      End If
End Function




