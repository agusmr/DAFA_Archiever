Attribute VB_Name = "ModBrowser"
' API untuk Browse for Folder
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBROWSEINFOTYPE As BROWSEINFOTYPE) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As Long)
Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
Private Declare Function LocalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal uBytes As Long) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long

Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsA" (ByVal pszPath As String) As Long

'Konstanta BFF

Private Const WM_USER = &H400
Private Const BFFM_SETSELECTIONA As Long = (WM_USER + 102)
Private Const BFFM_SETSELECTIONW As Long = (WM_USER + 103)
Private Const LPTR = (&H0 Or &H40)

' Type for Browse For Folder
Private Type BROWSEINFOTYPE
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type
' Fungsi Untuk menangani BFF
' ==========================================================================================================================================
Private Function BrowseCallbackProcStr(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    If uMsg = 1 Then
        Call SendMessage(hWnd, BFFM_SETSELECTIONA, True, ByVal lpData)
    End If
End Function
Private Function FunctionPointer(FunctionAddress As Long) As Long
    FunctionPointer = FunctionAddress
End Function
Public Function BrowseForFolder(frm As Form, PilihanPath As String, Judul As String) As String
    Dim Browse_for_folder As BROWSEINFOTYPE
    Dim itemID As Long
    Dim selectedPathPointer As Long
    Dim tmpPath As String * 256
With Browse_for_folder
    .hOwner = frm.hWnd ' Window Handle
    .lpszTitle = Judul ' Dialog Title
    .lpfn = FunctionPointer(AddressOf BrowseCallbackProcStr) ' Dialog callback function that preselectes the folder specified
    selectedPathPointer = LocalAlloc(LPTR, Len(PilihanPath) + 1) ' Allocate a string
    CopyMemory ByVal selectedPathPointer, ByVal PilihanPath, Len(PilihanPath) + 1 ' Copy the path to the string
    .lParam = selectedPathPointer ' The folder to preselect
End With
itemID = SHBrowseForFolder(Browse_for_folder) ' Execute the BrowseForFolder API
If itemID Then
    If SHGetPathFromIDList(itemID, tmpPath) Then ' Get the path for the selected folder in the dialog
        BrowseForFolder = Left$(tmpPath, InStr(tmpPath, vbNullChar) - 1) ' Take only the path without the nulls
    End If
    Call CoTaskMemFree(itemID) ' Free the itemID
End If
Call LocalFree(selectedPathPointer) ' Free the string from the memory
End Function
