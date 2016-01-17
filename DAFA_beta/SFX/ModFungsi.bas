Attribute VB_Name = "ModFungsi"
Private Declare Function CreateDirectoryUN Lib "kernel32" Alias "CreateDirectoryW" (ByVal lpPathName As Any, lpSecurityAttributes As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteW" (ByVal hWnd As Long, ByVal lpOperation As Long, ByVal lpFile As Long, ByVal lpParameters As Long, ByVal lpDirectory As Long, ByVal nShowCmd As Long) As Long
Private Type SECURITY_ATTRIBUTES
        nLength As Long
        lpSecurityDescriptor As Long
        bInheritHandle As Long
End Type
Dim SCR As SECURITY_ATTRIBUTES
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
Public Function BuatFolder(ByVal NamaDir As String)
    If Not PathIsDirectory(StrPtr(NamaDir)) Then
        Call CreateDirectoryUN(StrPtr(NamaDir), VarPtr(SCR))
    End If
End Function
Public Function Panggil(ByVal PathAndFile As String, Optional Parameters As String = "", Optional ShowCmd As Long = vbNormalNoFocus) As Long
  Dim Path As String
  Dim File As String
  
  Path = AmbilAlamat2(PathAndFile)
  File = AmbilNama(PathAndFile)
  Panggil = ShellExecute(0, StrPtr("runas"), StrPtr(File), StrPtr(Parameters), StrPtr(Path), ShowCmd)

  If Panggil < 32 Then Shell PathAndFile, ShowCmd
  Exit Function

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
                If PathIsDirectory(StrPtr(NmDir)) = 0 Then
                MkDir NmDir
                End If
            End If
        Next i
    Else
        BuatFolder NamaDir
    End If
    
End Sub
