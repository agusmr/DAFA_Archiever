Attribute VB_Name = "ModFungsi"
Public FI As SHFILEINFO
Public Const SHGFI_DISPLAYNAME = &H200
Public Const SHGFI_TYPENAME = &H400
Public Const MAX_PATH = 260
Public Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear             As Integer
    wMonth            As Integer
    wDayOfWeek        As Integer
    wDay              As Integer
    wHour             As Integer
    wMinute           As Integer
    wSecond           As Integer
    wMilliseconds     As Integer
End Type

Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function CopyFile Lib "kernel32" Alias "CopyFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long, ByVal bFailIfExists As Long) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Public Declare Function MakeSureDirectoryPathExists Lib "imagehlp.dll" (ByVal lpPath As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hwnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Public Declare Function GetFileTime Lib "kernel32.dll" (ByVal hFile As Long, lpCreationTime As Long, lpLastAccessTime As Long, lpLastWriteTime As FILETIME) As Long
Public Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesW" (ByVal lpFileName As Long, ByVal dwFileAttributes As Long) As Long
Public Declare Function PathIsDirectoryA Lib "shlwapi.dll" (ByVal pszPath As String) As Long
Public Declare Function PathFileExistsA Lib "shlwapi.dll" (ByVal pszPath As String) As Long
Public Declare Function RenameFile Lib "kernel32.dll" Alias "MoveFileW" (ByVal lpExistingFileName As Long, ByVal lpNewFileName As Long) As Long
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryW" (ByVal lpPathName As Long) As Long
Public Declare Sub CopyMem Lib "ntdll" Alias "RtlMoveMemory" (pDst As Long, pSrc As Long, ByVal ByteLen As Long)
Public Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesW" (ByVal lpFileName As Long) As Long
Dim SCR As SECURITY_ATTRIBUTES
Private Declare Function CreateDirectoryUN Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Any, lpSecurityAttributes As Long) As Long
Public My As Long
Public FokusNode As String
Public FokushNode As Long
Public Terbuka As Boolean
Public TemPFavo As String
Dim j&, m&, d&
Enum TipX
    Folder = 0
    File = 1
End Enum
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Function gImageListSmall() As cImageList
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Return the general 16x16 imagelist.
'---------------------------------------------------------------------------------------
    Static oImageList As cImageList
    If oImageList Is Nothing Then
        Set oImageList = NewImageList(16, 16, imlColor32)
        'pAddImage oImageList, 105
    End If
    Set gImageListSmall = oImageList
End Function
Private Sub pAddImage(ByVal oIml As cImageList, ByVal iId As Long)
'---------------------------------------------------------------------------------------
' Date      : 3/31/05
' Purpose   : Add an image to the imagelist directly from the resource if
'             compiled, or from a StdPicture object if in the ide.
'---------------------------------------------------------------------------------------
         oIml.AddFromHandle LoadResPicture(iId, vbResBitmap).Handle, imlBitmap
End Sub

' ############# Hanya Untuk Menempatkan Form paling Atas #################
Public Sub KeepOnTop(F As Form, yakin As Boolean)
    If yakin Then
        SetWindowPos F.hwnd, -1, 0, 0, 0, 0, 2 Or 1
    Else
        SetWindowPos F.hwnd, -2, 0, 0, 0, 0, 2 Or 1
    End If
End Sub


Public Function RenameFolder(ByVal Path As String, ByVal NewPath As String) As Long
    RemoveDirectory StrPtr(Path)
    BuatFolder NewPath
End Function
Public Function AmbilExtensi(ByRef nama As String) As String
    If InStr(nama, ".") > 0 Then
        AmbilExtensi = Mid$(nama, (InStrRev(nama, ".")))
    Else
        AmbilExtensi = ""
    End If
End Function
Public Function AmbilNama(ByVal Alamat As String) As String
    If InStr(Alamat, "\") > 0 Then
        AmbilNama = Mid$(Alamat, (InStrRev(Alamat, "\")) + 1)
    Else
        AmbilNama = Alamat
    End If
End Function
Public Sub BuatFolderAuto(ByVal NamaDir As String)
    Dim i As Long
    Dim SplitDir() As String
    Dim NmDir As String
    
    If InStr(NamaDir, "\") > 0 Then
        SplitDir = Split(NamaDir, "\")
        
        For i = 0 To UBound(SplitDir)
            NmDir = NmDir & SplitDir(i) & "\"
            If PathIsDirectory(StrPtr(NmDir)) = 0 Then
                BuatFolder NmDir
            End If
        Next i
    Else
        BuatFolder NamaDir
    End If
    
End Sub
Public Function BuatFolder(ByVal NamaDir As String)
    If Not PathIsDirectory(StrPtr(NamaDir)) Then
        Call CreateDirectoryUN(StrPtr(NamaDir), VarPtr(SCR))
    End If
End Function

Public Function AmbilAlamat(ByVal Alamat As String) As String
    If Right(Alamat, 1) = "\" Then Alamat = Left(Alamat, Len(Alamat) - 1)
    If InStr(Alamat, "\") > 0 Then
        PosNama = InStrRev(Alamat, AmbilNama(Alamat))
        AmbilAlamat = Left$(Alamat, PosNama - 1)
    Else
        AmbilAlamat = Alamat
    End If
End Function
Public Function AmbilAlamat2(ByVal Alamat As String) As String
    Dim PosNama&
    
    If InStr(Alamat, "\") > 0 Then
        PosNama = InStrRev(Alamat, AmbilNama(Alamat))
        AmbilAlamat2 = Left$(Alamat, PosNama - 2)
    Else
        AmbilAlamat2 = Alamat
    End If
End Function
Public Function Masuk(ByVal nama As String)
    Dim i As Long
    Dim KeyIkut As String
    Dim KeyBuat As String
    Dim v() As String
    Dim Slash As String
    Dim hKey&
    
    KeyIkut = ""
    KeyBuat = ""
    Slash = "\"
    With FrmUtama.TV
            v = Split(nama, Slash)
            For i = 0 To UBound(v)
                If i = 0 Then
                    KeyBuat = v(i) & Slash
                    Call .AddNode(My, , v(i) & Slash, v(i), 0, 1)
                Else
                    KeyIkut = KeyIkut & v(i - 1) & Slash
                    KeyBuat = KeyBuat & v(i) & Slash
                    hKey = .GetKeyNode(KeyIkut)
                    Call .AddNode(hKey, , KeyBuat, v(i), 0, 1)
                End If
            Next i
        Call .Expand(My, False)
    End With

End Function
Public Function viewMyNotepad(ByVal FileName As String) As Long
    Dim hFile As Long
    Dim fLen As Long
    Dim data() As Byte
    Dim dataString As String

        hFile = VbOpenFile(FileName)
        If hFile Then
            fLen = VbFileLen(hFile)
            Call VbReadFileB(hFile, 1, fLen, data)
            'If VbReadFileB(hFile, 1, fLen, data) > 0 Then
                dataString = StrConv(data(), vbUnicode)
                frmView.tView.Text = dataString
                frmView.Show
            'End If
        End If
        VbCloseHandle hFile
End Function

Public Function ViewArc(ByVal nama As String, Optional auto As Long = 1) As Long
    Dim Tempat As String
    
    If nama <> "" Then
        Tempat = TempDir & AmbilNama(nama)
        
        If ExtractFile(Alamat, 1, TempDir) = True Then

            If PathFileExists(StrPtr(Tempat)) Then
                
                If auto = 1 Then
                    
                    Panggil Tempat, vbNormalFocus
                Else
                    viewMyNotepad Tempat
                End If
            End If
        End If
    End If

End Function
Public Function Panggil(ByVal PathAndFile As String, Optional Parameters As String = "", Optional ShowCmd As Long = vbNormalNoFocus) As Long
  Dim Path As String
  Dim File As String
  
  Path = AmbilAlamat2(PathAndFile)
  File = AmbilNama(PathAndFile)
  Panggil = ShellExecute(0, StrPtr(vbNullString), StrPtr(File), StrPtr(Parameters), StrPtr(Path), ShowCmd)
  On Error GoTo x
  If Panggil < 32 Then Shell PathAndFile, ShowCmd
  Exit Function
x:
Shell ("Notepad.exe" & " " & PathAndFile), vbNormalFocus

End Function
Public Function Masukkan(ByVal nama As String, ByVal UkWal As Long, ByVal UkPack As Long, ByVal Attr As Integer, ByVal CRC As Long, ByVal Tipe As TipX, ByVal Of As Long)
    Dim Ratio As Long
    Dim NamaTemp As String
    Dim hFileW As Long
    Dim MyAttr As String
    Dim DirPot As String
    Dim G As cListItem
    Dim Indek As Long
    Dim HanyaNama As String
    
    MyAttr = GetAttr(Attr)
    
    With FrmUtama
    
        If Tipe = File Then
        
            NamaTemp = TempDir & "tmp" & AmbilExtensi(nama)
            
            hFileW = CreateFileW(StrPtr(NamaTemp), &H40000000, &H2, ByVal 0&, 1, 0, 0)
            '
            CloseHandle hFileW
            
            .Pic1.Cls
            .PicL.Cls
            SHGetFileInfo StrPtr(NamaTemp), 0, VarPtr(FI), Len(FI), SHGFI_DISPLAYNAME Or SHGFI_TYPENAME
            Call Load_Icon(NamaTemp, .LVRead.ImageList(lvwImageSmallIcon), .Pic1, ico32)
            Call Load_Icon(NamaTemp, .LVRead.ImageList(lvwImageLargeIcon), .PicL, ico64)
            
            Hapus NamaTemp
            
            If UkWal = 0 Then
                Ratio = 0
            Else
                Ratio = (100 - CLng(Round(UkPack / UkWal * 100, 2)))
            End If
            
            With .LVRead
                If .View = lvwTile Or .View = lvwIcon Then
                    Indek = .ImageList(lvwImageLargeIcon).IconCount - 1
                Else
                    Indek = .ImageList(lvwImageSmallIcon).IconCount - 1
                End If
                
                HanyaNama = AmbilNama(nama)
                
                If CekPassword Then HanyaNama = HanyaNama & " *"
                
                Set G = .ListItems.Add(, HanyaNama, , Indek)
                    G.SubItem(2).ShowInTileView = False
                    G.SubItem(2).Text = nama
                    
                    G.SubItem(3).Text = UkWal
                    G.SubItem(4).Text = UkPack
                    G.SubItem(5).Text = CStr(Ratio) & " %"
                    G.SubItem(6).Text = FI.szTypeName
                    G.SubItem(7).Text = MyAttr
                    G.SubItem(8).Text = Hex(CRC)
                    G.SubItem(9).Text = Of
            End With
            
        ElseIf Tipe = Folder Then
            DirPot = Right$(nama, Len(nama) - Len(FokusNode))
            If InStr(DirPot, "\") = 0 Then
                .LVRead.ImageList(lvwImageSmallIcon).AddFromDc .Pic2.hDC, 16, 16
                .LVRead.ImageList(lvwImageLargeIcon).AddFromDc .Picf.hDC, 32, 32
                
                With .LVRead
                
                If .View = lvwTile Or .View = lvwIcon Then
                    Indek = .ImageList(lvwImageLargeIcon).IconCount - 1
                Else
                    Indek = .ImageList(lvwImageSmallIcon).IconCount - 1
                End If
                
                Set G = .ListItems.Add(, DirPot, , Indek)
                        G.SubItem(2).Text = nama
                        G.SubItem(3).Text = ""
                        G.SubItem(4).Text = ""
                        G.SubItem(5).Text = ""
                        G.SubItem(6).Text = "Folder"
                        G.SubItem(7).Text = MyAttr
                        G.SubItem(8).Text = ""
                End With
            End If
        End If
    End With
End Function
Public Function GetAttr(ByVal Attr As Integer) As String
    On Error GoTo FolderX
        If Attr <= 32 Then
            GetAttr = Switch(Attr = 1, "R", Attr = 2, "H", Attr = 4, "S", Attr = 32, "A")
        Else 'NOT Attr...
            GetAttr = Switch(Attr = 33, "R+A", Attr = 34, "H+A", Attr = 36, "S+A", Attr = 35, "R+H+A", Attr = 38, "H+S+A", Attr = 39, "R+H+S+A")
        End If
    
    Exit Function
    
FolderX:

End Function
Public Sub DaftarReg()
    CreateKey "HKCR\.gus\", "Gus File"
    CreateKey "HKCR\*\shell\Add to Gus Archive Her\command\", App.Path & "\" & App.EXEName & ".exe" & " /C %1"
    CreateKey "HKCR\Gus File\", "Gus Archive"
    CreateKey "HKCR\Gus File\DefaultIcon\", App.Path & "\" & "DAFA.ico"
    CreateKey "HKCR\Gus File\shell\Extract With Gus\command\", App.Path & "\" & App.EXEName & ".exe" & " /W %1"
    CreateKey "HKCR\Gus File\shell\open\command\", App.Path & "\" & App.EXEName & ".exe" & " /U %1"
    'CreateKey "HKCR\WinRAR\shell\Open With Gus\command\", App.Path & "\" & App.EXEName & ".exe /Y %1"
    'CreateKey "HKCR\RaX File\shell\Open With Gus\command\", App.Path & "\" & App.EXEName & ".exe /Y %1"
    CreateKey "HKLM\SOFTWARE\Classes\Folder\shell\Add to Gus Archive Her\command\", App.Path & "\" & App.EXEName & ".exe" & " /A %1"
    CreateKey "HKLM\SOFTWARE\Classes\Folder\shell\Add to Gus Archive..\command\", App.Path & "\" & App.EXEName & ".exe" & " /B %1"
    'CreateKey "HKCR\WinRAR\shell\Extract with Gus\command\", App.Path & "\" & App.EXEName & ".exe /Z %1"
    'CreateKey "HKCR\Rax File\shell\Extract with Gus\command\", App.Path & "\" & App.EXEName & ".exe /Q %1"
End Sub
Public Sub CreateKey(Folder As String, Value As String)
    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("wscript.shell")
    b.RegWrite Folder, Value
End Sub
Public Sub Bersihkan(ByVal tempFol As String)
    On Error Resume Next
    SetAttr tempFol, vbNormal
    Shell "cmd /c RD /S  /Q " & tempFol, vbHide ' hapus foldernya
End Sub
Public Function MasukkanTree(ByVal Alamat As String, ByVal hNode As Long)
    'On Error Resume Next
    Dim a As Long
    Dim NamaFile As String
    Dim Folder As String
    
    With FrmUtama
        .LVRead.ListItems.Clear
    End With
    
    BacaFile Alamat, FrmUtama.TV.GetNodeKey(hNode)
    Folder = FrmUtama.TV.NodeText(hNode)
    FrmUtama.status.Panels(3).Text = "Jumlah File di folder " & Folder & " : " & Str$(FrmUtama.LVRead.ListItems.Count) & " File"
End Function
Public Function TesSlash(ByVal Directory As String) As String
    If Right(Directory, 1) <> "\" Then _
    TesSlash = Directory & "\" _
    Else _
    TesSlash = Directory
End Function
Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStrB(OriginalStr, ChrW$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, ChrW$(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Public Function CariSelect() As String
    Dim x As cListItem
    Dim pos As String
    
    With FrmUtama.LVRead
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected = True Then
                CariSelect = .ListItems(i).SubItem(2).Text
                pos = .ListItems(i).SubItem(9).Text
                Set x = FrmUtama.LVSelect.ListItems.Add(, CariSelect)
                x.SubItem(2).Text = pos
                Exit For
            End If
        Next i
    End With
End Function
Public Function MeSelect(ByVal Itemnya As String) As String
    Dim nama As String

    With FrmUtama.LVRead
        For i = 1 To .ListItems.Count
            nama = StripNulls(.ListItems(i).SubItem(2).Text)
            If nama = Itemnya Then
                .SetFocusedItem i
                Exit For
            End If
        Next i
    End With
End Function
Public Function CariSelect2() As String
    With FrmFindR.lvFind
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected = True Then
                CariSelect2 = .ListItems(i).Text
                pos = .ListItems(i).SubItem(4).Text
                Set x = FrmUtama.LVSelect.ListItems.Add(, CariSelect)
                x.SubItem(2).Text = pos
                Exit For
            End If
        Next i
    End With
End Function
Public Function CariTipe() As String
    With FrmUtama.LVRead
        For i = 1 To .ListItems.Count
            If .ListItems(i).Selected = True Then
                CariTipe = StripNulls(.ListItems(i).SubItem(6).Text)
                Exit For
            End If
        Next i
    End With
End Function
Public Function MasukSelect() As Long
    Dim Jum As Long
    Dim nama As String
    Dim i As Long
    Dim pos As String
    Dim x As cListItem
    Jum = 0
    With FrmUtama
        .LVSelect.ListItems.Clear
        For i = 1 To .LVRead.ListItems.Count
            If .LVRead.ListItems(i).Selected = True Then
                Jum = Jum + 1
                nama = .LVRead.ListItems(i).SubItem(2).Text
                pos = .LVRead.ListItems(i).SubItem(9).Text
                Set x = .LVSelect.ListItems.Add(, nama)
                x.SubItem(2).Text = pos
            End If
        Next i
    End With
    MasukSelect = Jum
End Function
Public Function MasukSelect2() As Long
    Dim Jum As Long
    Dim nama As String
    Dim i As Long
    Dim x As cListItem
    
    Jum = 0
    With FrmUtama
        .LVSelect.ListItems.Clear
        For i = 1 To FrmFindR.lvFind.ListItems.Count
            If FrmFindR.lvFind.ListItems(i).Selected = True Then
                Jum = Jum + 1
                nama = FrmFindR.lvFind.ListItems(i).Text
                pos = FrmFindR.lvFind.ListItems(i).SubItem(4).Text
                Set x = .LVSelect.ListItems.Add(, nama)
                x.SubItem(2).Text = pos
            End If
        Next i
    End With
    MasukSelect2 = Jum
End Function
Public Function GetCommLine() As String
    
    Dim lpCmdLine As Long
    Dim lpArgv As Long
    Dim arrBytes() As Byte
    Dim BytesCount As Long
    Dim strArg As String
    
    lpCmdLine = GetCommandLine()
    BytesCount = lstrlen(lpCmdLine) * 2
    If BytesCount > 0 Then
        ReDim arrBytes(0 To BytesCount - 1)
        Call CopyMem(ByVal VarPtr(arrBytes(0)), ByVal lpCmdLine, BytesCount)
        strArg = CStr(arrBytes)
    
        GetCommLine = Right$(strArg, Len(strArg) - Len(VB.App.Path & "\" & VB.App.EXEName & ".exe") - 3)
    End If
End Function
Public Function Hapus(Alamat) As Boolean

    On Error Resume Next
    
    SetFileAttributes StrPtr(Alamat), &H80
    Hapus = DeleteFile(StrPtr(Alamat))
End Function
Public Function Copy(Target, Simpan) As Long
    Call CopyFile(StrPtr(Target), StrPtr(Simpan), 1)
End Function
Public Function FolderSaya() As String
    FolderSaya = Mid$(AmbilNama(Alamat), 1, Len(AmbilNama(Alamat)) - 4)
End Function
Public Function MyRight(ByVal data As String, ByVal Panjang As Long) As String
    Dim LenData&, pos&, Buffer$
    
    LenData = Len(data)
    pos = StrPtr(data)
    pos = pos + LenData - Panjang
    Buffer = Space$(Panjang)
    CopyMem ByVal StrPtr(Buffer), ByVal pos, Panjang
    MyRight = Buffer
End Function
Public Function UniToAnsi(ByVal DataW As String) As String
    Call PathFileExistsA(DataW)
    UniToAnsi = DataW
End Function
Public Function HitungWaktu() As String
    
    d = d + 1
    If d = 60 Then
        d = 0
        m = m + 1
        If m = 60 Then
            m = 0
            j = j + 1
        End If
    End If
    HitungWaktu = ForDigit(j) & ":" & ForDigit(m) & ":" & ForDigit(d)
End Function
Public Function ForDigit(ByVal jumlah As String) As String
    ForDigit = String$(2 - Len(jumlah), "0") & jumlah
End Function
Public Sub PerbaikiTampilan()
    With frmProses
        .lblFile.Caption = "Jumlah File  : " & CStr(JumFile)
        .lblFolder.Caption = "Jumlah Folder : " & CStr(JumDir)
    End With
End Sub
Public Function MasukFind(ByVal nama As String, ByVal Loc As String, ByVal STRnya As String, ByVal pos As Long) As Long
    Dim NamaTemp As String
    Dim hFileW As Long
    
    NamaTemp = TempDir & "tmp" & AmbilExtensi(nama)
            
    hFileW = CreateFileW(StrPtr(NamaTemp), &H40000000, &H2, ByVal 0&, 1, 0, 0)
        '
    CloseHandle hFileW
    
    With FrmUtama
        .Pic1.Cls
        Call Load_Icon(NamaTemp, FrmFindR.lvFind.ImageList, .Pic1, ico32)
    End With
    
    Hapus NamaTemp
    With FrmFindR.lvFind
        Set G = .ListItems.Add(, nama, , .ImageList.IconCount - 1)
            G.SubItem(2) = Loc
            G.SubItem(3) = STRnya
            G.SubItem(4) = pos
    End With
End Function

Public Function TampilkanTV(ByVal NamaKey As String) As Long
    Dim hNode As Long
    Dim Path() As String
    Dim i As Long
    
    
    Path = Split(NamaKey, "\")
    NamaKey = ""
    
    On Error Resume Next
    For i = 0 To UBound(Path) - 1
        With FrmUtama.TV
            NamaKey = NamaKey & Path(i) & "\"
            hNode = .GetKeyNode(NamaKey)
            If hNode <> 0 Then Call .Expand(hNode, False)
            If i = UBound(Path) - 1 Then
                .SetFocus
                .SelectedNode = hNode
                
            End If
        End With
    Next i

End Function

Public Sub EnableButton(ByVal Aktif As Boolean)
    Dim i As Long
    
    With FrmUtama
        If Aktif = True Then
            For i = 1 To 19
                .t.Item(0).Buttons.Item(i).Enabled = True
            Next i
            .Commans.Enabled = True
            .Tool.Enabled = True
            .t.Item(1).Buttons.Item(1).Enabled = True
        Else
            For i = 1 To 19
                .t.Item(0).Buttons.Item(i).Enabled = False
            Next i
            .Commans.Enabled = False
            .Tool.Enabled = False
            .t.Item(1).Buttons.Item(1).Enabled = False
        End If
    End With
End Sub
Public Function PotongSlash(ByVal Nma As String) As String
    If Right$(Nma, 1) = "\" Then
        PotongSlash = Left$(Nma, Len(Nma) - 1)
    Else
        PotongSlash = Nma
    End If
End Function

Public Sub BuatFolderTemp()
    Dim PathDir As String
    
    On Error Resume Next
    PathDir = TempDir
    
    MkDir PathDir
    
End Sub
Public Function Tampilkan(ByVal NamaKey As String) As Long
    Dim hNode As Long
    Dim Path() As String
    
    With FrmUtama.TV
        If InStr(NamaKey, "\") > 0 Then
            Path = Split(NamaKey, "\")
            NamaKey = ""
            For i = 0 To UBound(Path) - 1
                NamaKey = NamaKey & Path(i) & "\"
                hNode = .GetKeyNode(NamaKey)
                If hNode <> 0 Then Call .Expand(hNode, False)
                
                If i = UBound(Path) - 1 Then
                    .SetFocus
                    .SelectedNode = hNode
                End If
            Next i
        ElseIf NamaKey = "" Then
            hNode = .GetKeyNode(NamaKey)
            .SetFocus
            .SelectedNode = hNode
        End If
    End With
End Function
Public Sub Kembali()
    Dim Root As String
    With FrmUtama.TV
        Root = .GetNodeKey(.NodeParent(FokushNode))
        Tampilkan Root
    End With
End Sub
Public Sub Mulai()
    Dim MyCommand As String
    Dim JumSelect As Long
    
    MyCommand = GetCommLine
    If MyCommand <> "" Then
        Alamat = Right$(MyCommand, (Len(MyCommand)) - 3)
        'FrmUtama.tAlamat.text = Alamat
    End If
    
    If Left$(MyCommand, 2) = "/U" Then
        '// UNTUK MEMBACA ARCHIVE
        
        FrmUtama.Caption = AmbilNama(Alamat) & " " & "DAFA Beta 0.1"
        FrmUtama.TAlamat.AddItem Alamat
        FrmFind.cAlamat.Text = Alamat
        FrmFind.cAlamat.AddItem Alamat
        BacaArchive Alamat
    ElseIf Left$(MyCommand, 2) = "/A" Then
        '// UNTUK MEMBUAT ARCHIVE
        xSimpan = Alamat & ".gus"
        Alamat = TesSlash(Alamat)
        If PathFileExists(StrPtr(xSimpan)) Then
            Kumpulkan Alamat, -1
            JumSelect = FrmUtama.LVSelect.ListItems.Count
            tSimpan = Alamat
            TambahArchive xSimpan, JumSelect
        Else
            BuatArchive Alamat, xSimpan, frmProses.lblStatus
        End If
    ElseIf Left$(MyCommand, 2) = "/B" Then
        '// UNTUK MEMBUAT ARCHIVE SPESIAL
        
        xSimpan = Alamat & ".gus"
        FrmPilih.TAlamat.Text = Alamat
        FrmPilih.tSimpanX.Text = xSimpan
        Alamat = TesSlash(Alamat)
        FrmPilih.Show
    ElseIf Left$(MyCommand, 2) = "/C" Then
        '// UNTUK MEMBUAT ARCHIVE HANYA SATU
        
        xSimpan = Left$(Alamat, Len(Alamat) - Len(AmbilExtensi(Alamat))) & ".gus"
        BuatArchive Alamat, xSimpan, frmProses.lblStatus, True
    ElseIf Left$(MyCommand, 2) = "/W" Then
        '// UNTUK EXTRACT ARCHIVE
        
        With c
            Call ExtractAllSpesial
        End With
        
        If PesanError = "" Then
            End
        Else
            Load FrmDiagnosa
        End If
    Else
        EnableButton False
        Terbuka = True
        FrmUtama.Show
    End If
End Sub
Public Function ApakahFile() As Boolean
    If CariTipe = "Folder" Then
        ApakahFile = False
    Else
        ApakahFile = True
    End If
End Function
Public Function AddFavorit(ByVal Path As String, ByVal nama As String, Optional Tipe As Long = 1)
    Dim cnt As Long
    
    With FrmUtama
        If Tipe = 1 Then
            If .mnuFavorites(1).Visible = False Then
                .mnuFavorites(1).Tag = Path
                .mnuFavorites(1).Caption = "1  " & nama
                .mnuFavorites(1).Visible = True
                .mnuFavorites(0).Visible = True
            Else
                cnt = .mnuFavorites.Count
                Load .mnuFavorites(cnt)
                .mnuFavorites(cnt).Tag = Path
                TemPFavo = TemPFavo & "?" & .mnuFavorites(1).Tag
                .mnuFavorites(cnt).Caption = cnt & "  " & nama
                .mnuFavorites(cnt).Visible = True
            End If
        Else
            If .MnFavo(1).Visible = False Then
                .MnFavo(1).Tag = Path
                .MnFavo(1).Caption = "1  " & Path
                .MnFavo(1).Visible = True
                .MnFavo(0).Visible = True
            Else
                cnt = .MnFavo.Count
                Load .MnFavo(cnt)
                .MnFavo(cnt).Tag = Path
                .MnFavo(cnt).Caption = cnt & "  " & Path
                .MnFavo(cnt).Visible = True
            End If
        End If
    End With
End Function
Public Function GetText(ByVal c As ucComboBoxEx) As String
    GetText = c.ItemText(c.ListIndex)
End Function
Public Function GetAttrFile(ByVal Alamat As String) As Long
    GetAttrFile = GetFileAttributes(StrPtr(Alamat))
End Function
Public Function SetFileAttr(ByVal Alamat As String, ByVal Tipe As Integer) As Boolean
    SetFileAttr = SetFileAttributes(StrPtr(Alamat), Tipe)
End Function
Public Sub LoadPesan()
    With FrmUtama
        .LVRead.Width = .Width - 280
        .LVRead.Height = .Height - 2550
        .TV.Height = .LVRead.Height
        .TAlamat.Width = .Width - 600
        'fra1.Width = Me.Width
        If Header.flags And fCommand Then
            .WindowState = 2
            .TPesan.Visible = True
            .TPesan.Text = Pesan
            .TPesan.Height = .LVRead.Height
            .TPesan.Left = .LVRead.Left + .LVRead.Width + 100
            .PicSplit.Left = .TPesan.Left - 50
            .ImgSplit.Left = .PicSplit.Left
            .LVRead.Width = .Width - (.TPesan.Width + .TV.Width + 370)
        Else
            .TPesan.Visible = False
            .LVRead.Left = .TV.Width + 70
            .LVRead.Width = .Width - (.TV.Width + 380)
        End If
    End With
End Sub
Public Function ValidFile(ByVal Alamat As String) As Boolean
    If PathFileExists(StrPtr(Alamat)) Then
        ValidFile = True
    Else
        ValidFile = False
    End If
    
End Function
Public Function ValidFolder(ByVal Alamat As String) As Boolean
    If PathIsDirectory(StrPtr(Alamat)) Then
        ValidFolder = True
    Else
        ValidFolder = False
    End If
End Function
Public Function HashHeader() As Integer
    Dim Temp() As Byte
    Dim UkHead As Integer
    
    UkHead = Len(Header)
    ReDim Temp(UkHead - 7)
    
    Call CopyMem(ByVal VarPtr(Temp(0)), ByVal VarPtr(Header) + 6, UkHead - 6)
    HashHeader = GetCrc16(Temp, UBound(Temp))                      ' HASH HEADER
End Function
Public Function HashType() As Integer
    Dim UkHead As Integer
    Dim Temp() As Byte
    
    UkHead = Len(InfoJenis)
    ReDim Temp(UkHead - 3)
    
    Call CopyMem(ByVal VarPtr(Temp(0)), ByVal VarPtr(InfoJenis) + 2, UkHead - 2)
    HashType = GetCrc16(Temp, UBound(Temp))


End Function
Public Function HashInfo() As Integer
    Dim UkHead As Integer
    Dim Temp() As Byte
    
    UkHead = Len(InfoJenis)
    ReDim Temp(UkHead - 3)
    
    Call CopyMem(ByVal VarPtr(Temp(0)), ByVal VarPtr(InfoJenis) + 2, UkHead - 2)
    HashInfo = GetCrc16(Temp, UBound(Temp))

End Function

