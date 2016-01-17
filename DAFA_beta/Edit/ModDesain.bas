VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStoreDc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Const DIB_RGB_COLORS                    As Long = &H0
Private Const BI_RGB                            As Long = &H0
Private Const BI_RLE4                           As Long = &H2
Private Const BI_RLE8                           As Long = &H1

Private Type SAFEARRAYBOUND
    cElements                                   As Long
    lLbound                                     As Long
End Type

Private Type SAFEARRAYID
    cDims                                       As Integer
    fFeatures                                   As Integer
    cbElements                                  As Long
    cLocks                                      As Long
    pvData                                      As Long
    Bounds                                      As SAFEARRAYBOUND
End Type

Private Type SAFEARRAY2D
    cDims                                       As Integer
    fFeatures                                   As Integer
    cbElements                                  As Long
    cLocks                                      As Long
    pvData                                      As Long
    Bounds(0 To 1)                              As SAFEARRAYBOUND
End Type

Private Type BITMAP
    bmType                                      As Long
    bmWidth                                     As Long
    bmHeight                                    As Long
    bmWidthBytes                                As Long
    bmPlanes                                    As Integer
    bmBitsPixel                                 As Integer
    bmBits                                      As Long
End Type

Private Type BITMAPINFOHEADER
    biSize                                      As Long
    biWidth                                     As Long
    biHeight                                    As Long
    biPlanes                                    As Integer
    biBitCount                                  As Integer
    biCompression                               As Long
    biSizeImage                                 As Long
    biXPelsPerMeter                             As Long
    biYPelsPerMeter                             As Long
    biClrUsed                                   As Long
    biClrImportant  As Long
End Type

Private Type RGBQUAD
    rgbBlue                                     As Byte
    rgbGreen                                    As Byte
    rgbRed                                      As Byte
    rgbReserved                                 As Byte
End Type

Private Type BITMAPINFO
    bmiHeader                                   As BITMAPINFOHEADER
    bmiColors                                   As RGBQUAD
End Type

Private Type GUID
    Data1                                       As Long
    Data2                                       As Integer
    Data3                                       As Integer
    Data4(7)                                    As Byte
End Type


Private Type PICTUREINFO
    Size                                        As Long
    Type                                        As Long
    hBmp                                        As Long
    hPal                                        As Long
    Reserved                                    As Long
End Type


Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpDst As Any, _
                                                                     lpSrc As Any, _
                                                                     ByVal Length As Long)

Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (lpDst As Any, _
                                                                     ByVal Length As Long)

Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, _
                                                                       lpDeviceName As Any, _
                                                                       lpOutput As Any, _
                                                                       lpInitData As Any) As Long

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, _
                                                             ByVal nWidth As Long, _
                                                             ByVal nHeight As Long) As Long

Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, _
                                                   ByVal hObject As Long) As Long

Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, _
                                             ByVal X As Long, _
                                             ByVal Y As Long, _
                                             ByVal nWidth As L