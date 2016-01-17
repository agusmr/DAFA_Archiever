VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.UserControl DirCtl 
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2175
   ScaleHeight     =   1605
   ScaleWidth      =   2175
   Begin VB.PictureBox Picture1 
      Height          =   1095
      Left            =   510
      ScaleHeight     =   1035
      ScaleWidth      =   1365
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   1425
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   12
         Left            =   30
         Picture         =   "DirCtl.ctx":0000
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   15
         Top             =   810
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   11
         Left            =   780
         Picture         =   "DirCtl.ctx":058A
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   14
         Top             =   540
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   9
         Left            =   300
         Picture         =   "DirCtl.ctx":0B14
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   13
         Top             =   540
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   8
         Left            =   60
         Picture         =   "DirCtl.ctx":109E
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   12
         Top             =   540
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   6
         Left            =   540
         Picture         =   "DirCtl.ctx":1628
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   11
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   5
         Left            =   270
         Picture         =   "DirCtl.ctx":1BB2
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   10
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   4
         Left            =   30
         Picture         =   "DirCtl.ctx":213C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   9
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   3
         Left            =   750
         Picture         =   "DirCtl.ctx":26C6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   8
         Top             =   30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   2
         Left            =   510
         Picture         =   "DirCtl.ctx":2C50
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   7
         Top             =   30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   1
         Left            =   270
         Picture         =   "DirCtl.ctx":31DA
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   6
         Top             =   30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "DirCtl.ctx":3764
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   5
         Top             =   30
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   7
         Left            =   750
         Picture         =   "DirCtl.ctx":3CEE
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   4
         Top             =   270
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   10
         Left            =   570
         Picture         =   "DirCtl.ctx":4278
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   13
         Left            =   420
         Picture         =   "DirCtl.ctx":4802
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   2
         Top             =   840
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.PictureBox picBuffer 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   14
         Left            =   1230
         Picture         =   "DirCtl.ctx":4D8C
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   1
         Top             =   630
         Visible         =   0   'False
         Width           =   240
      End
   End
   Begin DAFA_Component.ucTreeView DirTree 
      Height          =   1575
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   2778
   End
End
Attribute VB_Name = "DirCtl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Public Event DirPath(ByVal spath As String)
Private Declare Function SHGetSpecialFolderPath Lib "shell32.dll" Alias "SHGetSpecialFolderPathA" (ByVal hWnd As Long, ByVal pszPath As String, ByVal CSIDL As Long, ByVal fCreate As Boolean) As Boolean
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeW" (ByVal nDrive As Long) As Long
Private Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long
Private Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As Long) As Long
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
Private Wfd As WIN32_FIND_DATA
Private FileSpec As String
Private UseFileSpec As Boolean
Private hFindFile As Long
Private Const DOT1 As String = "."
Private Const DOT2 As String = ".."
Private Const Bintang As String = "*"
Private Const Slas As String = "\"
Private Const Semua As String = "*.*"
Private Type WIN32_FIND_DATA
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

Public Enum IDFolder
    DESKTOP = &H0
    PROGRAMS = &H2
    Controls = &H3
    Printers = &H4
    PERSONAL = &H5
    FAVORITES = &H6
    STARTUP = &H7
    RECENT = &H8
    SENDTO = &H9
    BITBUCKET = &HA
    STARTMENU = &HB
    DESKTOPDIRECTORY = &H10
    DRIVES = &H11
    NETWORK = &H12
    NETHOOD = &H13
    Fonts = &H14
    TEMPLATES = &H15
    ALL_USER_STARTUP = &H18
    DEKSTOP_PATH = &H19
    WINDOWS_DIR = &H24
    SYSTEM_DIR = &H25
    PROGRAM_FILE = &H26
End Enum

Dim My&, Sys&, Spec&, sDir&, p&
Public Sub OutPutPath(lst As Collection)
    With DirTree
        If .NodeChecked(My) = False Then Exit Sub
    End With
End Sub

Public Function StripNulls(ByVal OriginalStr As String) As String
    If (InStrB(OriginalStr, ChrW$(0)) > 0) Then
        OriginalStr = Left$(OriginalStr, InStr(OriginalStr, ChrW$(0)) - 1)
    End If
    StripNulls = OriginalStr
End Function
Public Sub ScanDir(ByVal Path As String, ByVal hFol As Long)
    Dim dirs As Long, dirbuff() As String, i As Long
    Dim Alamat As String
    Dim NameDir As String
        hFindFile = FindFirstFileW(StrPtr(Path & Bintang), VarPtr(Wfd))
        If hFindFile <> INVALID_HANDLE_VALUE Then
            Do
                If Wfd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY Then
                    If AscW(Wfd.cFileName) <> vbKeyDot Then
                        NameDir = StripNulls(Wfd.cFileName)
                        If LCase(NameDir) = "my music" Then
                            Call DirTree.AddNode(hFol, , Path & NameDir & "\", NameDir, 14, 14)
                        ElseIf LCase(NameDir) = "my pictures" Then
                            Call DirTree.AddNode(hFol, , Path & NameDir & "\", NameDir, 13, 13)
                        Else
                            Call DirTree.AddNode(hFol, , Path & NameDir & "\", NameDir, 0, 1)
                        End If
                    End If
                End If
            Loop While FindNextFileW(hFindFile, VarPtr(Wfd))
            Call FindClose(hFindFile)
        End If
Ex:
Exit Sub
    Call FindClose(hFindFile)
End Sub

Private Sub pvInitializeTreeView1()
    
  Dim i As Long
    
    With DirTree
        
        Call .Initialize
        Call .InitializeImageList
        
        For i = 0 To 14
            Call .AddIcon(picBuffer(i).Picture)
        Next i
        .ItemHeight = 18
        .HasButtons = True
        .HasLines = True
        .HasRootLines = True
        .TrackSelect = True
        .LabelEdit = True
        .BackColor = vbWhite
        .ForeColor = &H0
        .LineColor = vbBlack
                
    End With
End Sub
Public Sub LoadTreeView()
    Dim DriveNum As String
    Dim DriveType As Long
    Dim MyIco As Byte
    Dim NameDrive As String
    Dim RootDrive As String
    Dim PathDoc As String
    Dim PathDekstop As String
    Dim PathShared As String
    Dim MyDoc As Long
    Dim MyShared As Long
    
    DriveNum = 64
    'On Error Resume Next
    With DirTree
        .Clear
        PathDekstop = GetSpecFolder(DESKTOPDIRECTORY) & "\"
        PathDoc = GetSpecFolder(PERSONAL) & "\"
        PathShared = GetSpecFolder(DEKSTOP_PATH) & "\"
        
        Spec = .AddNode(, rFirst, PathDekstop, "Desktop", 12, 12)
        MyDoc = .AddNode(Spec, , PathDoc, "My Documents", 10, 10)
        ScanDir PathDoc, MyDoc
        Call .Expand(Spec, False)
        Call .Expand(MyDoc, False)
        My = .AddNode(, , "My Computer", "My Computer", 5, 5)
            Do
                DriveNum = DriveNum + 1
                RootDrive = Chr$(DriveNum)
                DriveType = GetDriveType(StrPtr(RootDrive & ":\"))
                If DriveNum > 90 Then Exit Do
                NameDrive = StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase)
                Select Case DriveType
                    Case 0
                        sDir = .AddNode(My, , RootDrive & ":\", NameDrive & " (" & RootDrive & ":)")
                        ScanDir RootDrive & ":\", sDir
                    Case 2
                        sDir = .AddNode(My, , RootDrive & ":\", "(" & RootDrive & ":)", 2, 2)
                        ScanDir RootDrive & ":\", sDir
                    Case 3
                        If NameDrive = "" Then NameDrive = "Local Disk"
                        sDir = .AddNode(My, , RootDrive & ":\", NameDrive & " (" & RootDrive & ":)", 4, 4)
                        ScanDir RootDrive & ":\", sDir
                    Case 4
                        sDir = .AddNode(My, , RootDrive & ":\", NameDrive & " (" & RootDrive & ":)")
                        ScanDir RootDrive & ":\", sDir
                    Case 5
                        If NameDrive = "" Then NameDrive = "CD/DVD-Drive"
                        sDir = .AddNode(My, , RootDrive & ":\", NameDrive & " (" & RootDrive & ":)", 3, 3)
                        ScanDir RootDrive & ":\", sDir
                    Case 6
                        sDir = .AddNode(My, , RootDrive & ":\", NameDrive & " (" & RootDrive & ":)")
                        ScanDir RootDrive & ":\", sDir
                End Select
            Loop
        Call .Expand(My, False)
    End With

End Sub


Private Sub DirTree_AfterLabelEdit(ByVal hNode As Long, Cancel As Integer, NewString As String)
    Dim NamaLawas As String
    Dim PotPath As String
    
    
   ' RenameFolder DirTree.GetNodeKey(hNode), NewString

End Sub

Private Sub DirTree_BeforeExpand(ByVal hNode As Long, ByVal ExpandedOnce As Boolean)
    
    Dim i As Integer, Nodeku As Long, Path As String
    
    With DirTree
        Nodeku = .NodeChild(hNode)
        For i = 1 To .NodeChildren(hNode)
            If i > 1 Then _
                Nodeku = .NodeNextSibling(Nodeku)
            If .NodeChildren(Nodeku) <> 0 Then _
                Exit Sub
            Path = .GetNodeKey(Nodeku)
            ScanDir Path, Nodeku
            
        Next
    End With
End Sub

Private Sub DirTree_NodeClick(ByVal hNode As Long)
    Dim Path As String
    
    Path = DirTree.GetNodeKey(hNode)
    ScanDir Path, hNode
    RaiseEvent DirPath(Path)
End Sub

Private Sub UserControl_Initialize()
    Call pvInitializeTreeView1
End Sub

Private Sub UserControl_Resize()
    With DirTree
        .Height = UserControl.Height
        .Width = UserControl.Width
    End With
End Sub

Private Function FixBuffer(ByVal sBuffer As String) As String

Dim NullPos As Long
    
    NullPos = InStr(sBuffer, Chr$(0))
    
    If NullPos > 0 Then
        FixBuffer = Left$(sBuffer, NullPos - 1)
    End If
    
End Function
'"C:\Documents and Settings\Administrator\My Documents"

' Fungsi Utama daptakan Path Folder Spesial
Public Function GetSpecFolder(ByVal lpCSIDL As IDFolder) As String

Dim spath As String
Dim lRet As Long
    
    spath = String$(255, 0)
    
    lRet = SHGetSpecialFolderPath(0&, spath, lpCSIDL, False)
    
    If lRet <> 0 Then
        GetSpecFolder = FixBuffer(spath)
    End If
    
End Function

Public Function Tampilkan(ByVal NamaKey As String) As Long
    Dim hNode As Long
    Dim Path() As String
    Dim i As Long
    
             On Error Resume Next
   
    Path = Split(NamaKey, "\")
    NamaKey = ""
    For i = 0 To UBound(Path) - 1
        NamaKey = NamaKey & TesSlash(Path(i))
        hNode = DirTree.GetKeyNode(NamaKey)
        If hNode <> 0 Then Call DirTree.Expand(hNode, False)
        If i = UBound(Path) - 1 Then
            DirTree.SetFocus
            DirTree.SelectedNode = hNode
            
        End If
    Next i

End Function



