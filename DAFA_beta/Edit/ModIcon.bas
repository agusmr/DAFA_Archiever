Attribute VB_Name = "ModIcon"
Public Const MAX_PATH As Long = 260
Private Const SHGFI_DISPLAYNAME = &H200, SHGFI_EXETYPE = &H2000, SHGFI_SYSICONINDEX = &H4000, SHGFI_LARGEICON = &H0, SHGFI_SMALLICON = &H1, SHGFI_SHELLICONSIZE = &H4, SHGFI_TYPENAME = &H400, ILD_TRANSPARENT = &H1, BASIC_SHGFI_FLAGS = SHGFI_TYPENAME Or SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon As Long
    iIcon As Long
    dwAttributes As Long
    szDisplayName As String * MAX_PATH
    szTypeName As String * 80
End Type

Public Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoW" (ByVal pszPath As Long, ByVal dwFileAttributes As Long, ByVal psfi As Long, ByVal cbSizeFileInfo As Long, ByVal uFlags As Long) As Long
Private Declare Function ImageList_Draw Lib "Comctl32.dll" (ByVal himl As Long, ByVal i As Long, ByVal hDCDest As Long, ByVal x As Long, ByVal y As Long, ByVal flags As Long) As Long
Private shinfo As SHFILEINFO, sshinfo As SHFILEINFO
Private Declare Function DrawIconEx Lib "user32.dll" (ByVal hDC As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Enum IconMode
    ico64 = &H0
    ico32 = &H1
End Enum
Public Sub RetrieveIcon(fName As String, DC As PictureBox, ukuRan As IconMode)
    Dim hImgSmall As Long
    Dim hImgLarge As Long
    
    hImgSmall = SHGetFileInfo(StrPtr(fName), 0&, VarPtr(shinfo), Len(shinfo), BASIC_SHGFI_FLAGS Or ukuRan)
    ImageList_Draw hImgSmall, shinfo.iIcon, DC.hDC, 0, 0, ILD_TRANSPARENT
End Sub
Public Function Load_Icon(FileName As String, iList As cImageList, PictureBox As PictureBox, TypeIcon As IconMode) As Long
    Dim SmallIcon As Long
    Dim IconIndex As Integer
    
    RetrieveIcon FileName, PictureBox, TypeIcon
    If TypeIcon = ico32 Then
        iList.AddFromDc PictureBox.hDC, 16, 16
    Else
        iList.AddFromDc PictureBox.hDC, 32, 32
    End If
End Function
Public Function Ex_Icon(FileName As String, PictureBox As PictureBox, TypeIcon As IconMode) As Long
    Dim SmallIcon As Long
    Dim IconIndex As Integer
    
    RetrieveIcon FileName, PictureBox, TypeIcon
End Function


