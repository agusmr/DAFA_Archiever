As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type
Private Type MODULEENTRY32
    dwSize As Long
    th32ModuleID As Long
    th32ProcessID As Long
    GlblcntUsage As Long
    ProccntUsage As Long
    modBaseAddr As Long
    modBaseSize As Long
    hModule As Long
    szModule As String * 256
    szExePath As String * 260
End Type
Private Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnID As NOTIFYICONDATA) As Boolean
Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Function KillProcess(pProcessID As Long)
    TerminateProcess OpenProcess(2035711, 1, pProcessID), 0
    DoEvents
End Function
Private Function GetTheProcesses() As String
Dim pProcess  As PROCESSENTRY32  'deklarasi variable
Dim sSnapShot As Long
Dim rReturn   As Integer
    sSnapShot = CreateToolhelp32Snapshot(15, 0) 'setting variable
    pProcess.dwSize = Len(pProcess) '
    Process32First sSnapShot, pProcess 'proses pertama ([System Process])
    Do 'lakukan looping
        GetTheProcesses = GetTheProcesses & StripNulls(pProcess.szExeFile) & ParseMe & pProcess.th32ProcessID & ParseMe 'adds to the string variable the next process and ID
        rReturn = Process32Next(sSnapShot, pProcess)
        DoEvents 'pindahkan pengerjaan ke memori
    Loop While rReturn <> 0
    CloseHandle sSnapShot
End Function
Private Function PathByPID(pid As Long) As String
    Dim cbNeeded As Long
    Dim Modules(1 To 200) As Long
    Dim Ret As Long
    Dim ModuleName As String
    Dim nSize As Long
    Dim hProcess As Long
    hProcess = OpenProcess(1024 Or 16, 0, pid)
    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded)
        If Ret <> 0 Then
            ModuleName = Space(260)
            nSize = 500
            Ret = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
            PathByPID = Left(ModuleName, Ret)
        End If
    End If
    Ret = CloseHandle(hProcess)
    If PathByPID = "" Then
        PathByPID = "SYSTEM"
    End If
    If Left(PathByPID, 4) = "\??\" Then
        PathByPID = Mid(PathByPID, 5, Len(PathByPID))
        Exit Function
    End If
    If Left(PathByPID, 12) = "\SystemRoot\" Then
        PathByPID = WindowsDirectory & "\" & Mid(PathByPID, 13, Len(PathByPID))
        Exit Function
    End If
End Function
Public Function matikan(NamaFile As String) As Long
    On Error GoTo ErrHandle
    Dim uProcess As PROCESSENTRY32
    Dim lProc As Long, hProcSnap As Long
    Dim ExePath As String
    Dim hPID As Long, hExit As Long
    Dim i As Integer
    uProcess.dwSize = Len(uProcess)
    hProcSnap = CreateToolhelp32Snapshot(&H2, 0&)
    lProc = Process32First(hProcSnap, uProcess)
    Do While lProc
        i = InStr(1, uProcess.szExeFile, Chr$(0))
        ExePath = UCase$(Left$(uProcess.szExeFile, i - 1))
        If UCase$(AmbilNamaFile(ExePath)) = UCase$(NamaFile) Then
            hPID = OpenProcess(1&, -1&, uProcess.th32ProcessID)
            hExit = TerminateProcess(hPID, 0&)
            Call CloseHandle(hPID)
        End If
        lProc = Process32Next(hProcSnap, uProcess)
    Loop
    Call CloseHandle(hProcSnap)
    Exit Function
ErrHandle:
End Function
Public Function AmbilNamaFile(Alamat) As String
pa = Alamat
AmbilNamaFile = Mid(pa, (InStrRev(pa, "\")) + 1)
End Function
Public Function M31Pattern(Alamat) As String
Dim temp() As Byte
Dim X, X2 As Long
Dim Num, num2 As Long
ReDim temp(202 + 400) As Byte
On Error GoTo Keluar
 If FileLen(Alamat) >= 4520 Then ' Ukuran File standarnya 4520 Byte
      Open Alamat For Binary As #1
          isiHexFile = Space(LOF(1))
          Get #1, , isiHexFile
          Get #1, 3911, temp
      Close #1
      For Num = 1 To 202 'UBound(Temp)
          X = X + temp(Num) ^ 3
      Next
      For num2 = 1 To 202  'Ubound(Temp2)
          X2 = X2 + temp(num2 + 400) ^ 3 ' Ambil Data dari Pos 4311
      Next
 Else ' Kalo gak standar Pake ini ngeceknya
    If FileLen(Alamat) >= 610 Then ' Ukuran minimal dengan M31 Pattern adalah 610 an Byte
         Open Alamat For Binary As #1
             isiHexFile = Space(LOF(1))
             Get #1, , isiHexFile
             Get #1, , temp
          Close #1
         For Num = 1 To 202
             X = X + Asc(Mid(isiHexFile, Num + 1, 1)) ^ 3
         Next
         For num2 = 1 To 202
             X2 = X2 + Asc(Mid(isiHexFile, num2 + 401, 1)) ^ 3
       Next
    End If
 End If
M31Pattern = Hex(X) & Hex(X2)
Keluar:
End Function
Private Function calc_byte(path_Picture As String) As String
Dim binT() As Byte
Dim count, long_hash As Double
ReDim binT(FileLen(path_Picture)) As Byte
Open path_Picture For Binary As #1
    Get #1, , binT
Close #1
For count = 1 To UBound(binT)
    long_hash = long_hash + binT(count) ^ 2
Next
calc_byte = Hex(long_hash)
End Function
Public Function I_Code(Alamat) As String
Pic1.Cls
IconExist = ExtractIconEx(Alamat, 0, ByVal 0&, hIcon, 1)
If IconExist <= 0 Then
    IconExist = ExtractIconEx(Alamat, 0, hIcon, ByVal 0&, 1)
    If IconExist <= 0 Then GoTo bersih
End If
DrawIconEx Pic1.hDC, 0, 0, hIcon, 0, 0, 0, 0, &H3 '--> lihat ket
SavePicture Pic1.Image, App.Path & "\pic.tmp"
I_Code = calc_byte(App.Path & "\pic.tmp")
Kill App.Path & "\pic.tmp"
Exit Function
bersih:
I_Code = ""
End Function
Public Function LihatFolder(Komentar)
  Dim n As Integer
  Dim IDList As Long
  Dim Result As Long
  Dim ThePath As String
  Dim BI As BrowseInfo
  With BI
    .hWndOwner = GetActiveWindow()
    .lpszTitle = lstrcat(Komentar, "")
    .ulFlags = BIF_RETURNONLYFSDIRS
  End With
  IDList = SHBrowseForFolder(BI)
  If IDList Then
    ThePath = String$(MAX_PATH, 0)
    Result = SHGetPathFromIDList(IDList, ThePath)
    Call CoTaskMemFree(IDList)
    n = InStr(ThePath, vbNullChar)
    If n Then ThePath = Left$(ThePath, n - 1)
  End If
  LihatFolder = ThePath
End Function
Public Function LihatProses(lstbox)
  Dim cb As Long
  Dim cbNeeded As Long
  Dim NumElements As Long
  Dim ProcessIDs() As Long
  Dim cbNeeded2 As Long
  Dim NumElements2 As Long
  Dim Modules(1 To 99999) As Long
  Dim lRet As Long
  Dim ModuleName As String
  Dim nSize As Long
  Dim hProcess As Long
  Dim i As Long
  Dim sModName As String
  Dim iModDlls As Long
  Dim iProcesses As Integer
  Dim FillProcessListNT As Integer
  lstbox.Clear
  cb = 8
  cbNeeded = 96
  Do While cb <= cbNeeded
    cb = cb * 2
    ReDim ProcessIDs(cb / 4) As Long
    lRet = EnumProcesses(ProcessIDs(1), cb, cbNeeded)
  Loop
  NumElements = cbNeeded / 4
  For i = 1 To NumElements
    hProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, ProcessIDs(i))
    If hProcess Then
      lRet = EnumProcessModules(hProcess, Modules(1), 200, cbNeeded2)
      If lRet <> 0 Then
        ModuleName = Space(MAX_PATH)
        nSize = 500
        lRet = GetModuleFileNameExA(hProcess, Modules(1), ModuleName, nSize)
        sModName = Left$(ModuleName, lRet)
        lstbox.AddItem sModName
        iProcesses = iProcesses + 1
        iModDlls = 1
        Do
          iModDlls = iModDlls + 1
          ModuleName = Space(MAX_PATH)
          nSize = 500
          lRet = GetModuleFileNameExA(hProcess, Modules(iModDlls), ModuleName, nSize)
          sChildModName = Left$(ModuleName, lRet)
          If sChildModName = sModName Then Exit Do
        'If Trim(sChildModName) <> "" Then lstbox.AddItem sChildModName
        Loop
      End If
    Else
      FillProcessListNT = 0
    End If
    lRet = CloseHandle(hProcess)
  Next i
  FillProcessListNT = iProcesses
End Function
Public Function StripNulls(OriginalStr) As String
    If (InStr(OriginalStr, Chr(0)) > 0) Then
        OriginalStr = Left(OriginalStr, InStr(OriginalStr, Chr(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Private Sub UserControl_Initialize()
    UserControl.Height = 15 * 15
    UserControl.Width = 15 * 15
End Sub
Private Sub UserControl_Resize()
    UserControl.Height = 15 * 15
    UserControl.Width = 15 * 15
End Sub


                                                                                                                                                                                                                                                                                         