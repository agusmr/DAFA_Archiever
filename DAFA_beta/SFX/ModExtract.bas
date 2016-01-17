Attribute VB_Name = "ModExtract"
    '==================================================================
    '==================================================================
    '===  AgrotiX Packing Algorithm                                 ===
    '===  Proggramer Agus Minanur Rohman                            ===
    '===  Karya Proggramer Santri Pondok Pesantren Darul Ma'arif    ===
    '===  Data 23 februari 2010                                     ===
    '===  Payaman, Solokuro, Lamongan, Jawa Timur, Indonesia        ===
    '===  HP 085732446543                                           ===
    '===  E-Mail comp.agus@yahoo.com                                ===
    '===  Thanks To Codenesia.com                                   ===
    '==================================================================
    '==================================================================
    ' Penting [ Dilarang keras menghilangkan Nama Penulis ]

'========================================================================
'========================================================================
'====================== STRUCTURE AGROTIX PACKING =======================
'========================================================================
'======================= BY. AGUS MINANUR ROHMAN ========================
'========================================================================
'========================================================================
Private Type ArchiveInformation
    Rat         As Long
    Versi       As String
    TotFile     As String
    TotPack     As String
    TotSize     As String
    Ratio       As String
    SFX         As String
    Dictionary  As String
    Recovary    As String
    Verification As String
    Password    As String
    Lock        As String
    Command     As String
End Type

Private Type HeaderGus
    Signature           As Long     ' SIGNATURE AGROTIX 1
    Hash_Head           As Integer  ' HASH HEADER
    StructurArchive     As Byte     ' MODEL ARCHIVE
    ' 1 = Model Awal
    
    MethodCompress      As Byte     ' MODEL COMPRESSION
    ' 1 = Method Deflate
    ' 2 = Method Bzib2
    ' 3 = Method LZMA
    ' 4 = Method Aplib
    ' 5 = Method fLz
    flags               As Integer  ' BERISI INFO UTAMA
    ' Command  = H1
    ' Encript  = H2
    ' Lock     = H4
    ' Split    = H8
    ' Solid    = H10
    ' Compress = H20
    '            H40
    CrcEncript          As Integer  ' BERISI CRC PASSWORD ENCRIPT JIKA
    Reserved(5)         As Integer  ' CADANGAN UNTUK UPDATE
End Type

Private Type CommandArchive
    CommandCrc      As Integer      ' BERISI CRC KOMENTAR
    LenUnPack       As Integer      ' BERISI UKURAN KOMENTAR SEBELUM DI PACK
    LenPack         As Integer      ' BERISI UKURAN KOMENTAR SETELAH DI PACK
End Type

Private Type InfoType
    Type_Hash       As Integer      ' BERISI HASH <Time> Sampai <CrcPass>
    Time            As Long         ' BERISI INFORMASI WAKTU FILE
    sFlags          As Integer      ' BERISI INFORMASI TENTANG FILE
    ' Password = H1
    ' Unicode  = H2
    ' File     = H4
    ' Folder   = H8
    '          = H10
    '            H20
    '            H40
    UkNamaFile      As Integer      ' BERISI UKURAN NAMA FILE
    Atribut         As Integer      ' BERISI ATRIBUT FILE
End Type

'=========================================================
'======= DISINI ADA NAMA FILE BERFORMAT STRING ===========
'=========================================================

'============= INFORMASI TAMBAHAN UNTUK FILE =============
Private Type UntukFile
    OffsetData      As Long
    CRC32           As Long
    UkAwalData      As Long
    UkAhirData      As Long
    CrcPass         As Integer
End Type

'===========================================================================
'===========================================================================
Public Header       As HeaderGus
Public InfoCommand  As CommandArchive
Public InfoJenis    As InfoType
Public InfoFile     As UntukFile
'===========================================================================
'===========================================================================
Public Const fCommand = &H1
Public Const FPassword = &H1
Public Const FEncript = &H2
Public Const FUnicode = &H2
Public Const FLock = &H4
Public Const FSplit = &H8
Public Const FSolid = &H10
Public Const FCompress = &H20
'===========================================================================
'===========================================================================
Const FFile = &H4
Const FFolder = &H8
Const DAFASignature = &H3994D44
Const DAFAHeader = &H80
Const VersiArchive = &H1
Const VersiCompres = &H5
'===========================================================================
'===========================================================================
Const MetDeflate = &H1
Const MetBzib2 = &H2
Const MetLZMA = &H3
Const MetAPlib = &H4
Const MetfLz = &H5
'===========================================================================
'===========================================================================

Public iArc As ArchiveInformation
Public PesanError   As String
Public pass         As String
Public SetPass      As Boolean
Public DiEncrip     As Boolean
Public DiPassword   As Boolean
Public DiPesan      As Boolean
Public DiLock       As Boolean
Public JumFile      As Long
Public JumDir       As Long
Public Alamat       As String
Public tSimpan      As String
Public MySize       As Long
Public YestoAll     As Boolean
Public NotoAll      As Boolean
Public RenameAll    As Boolean
Public NameRename   As String
Public xSimpan      As String
Public Offset       As Long
Public pesan        As String
Public MethodC      As Byte
Public RatioC       As Long
Public hFileDibuat As Long
Public ErrorHard As Boolean
Public Deskripsi As String

Dim nama As String
Dim NamaE As String
Dim UkWal As Long
Dim UkPack As Long
Dim Attr As Integer
Dim hFileR As Long
Dim hFileW As Long
Dim ukuRan As Long
Dim isi() As Byte
Dim data() As Byte
Dim Hasil() As Byte
Dim ret As Long
Dim CRC As String
Dim NamaFolder As String
Dim Flg As Integer
Dim Fuokus As String
Dim TotPack As Long
Dim TotUk As Long
Dim Rat As Integer
Dim Pointer As Long
Dim NamaX As String
Dim TempD() As Byte
Dim z As Long
Dim x As Long
Dim hfileX As Long

Public Function ExtractArchive(ByVal Alamat As String, ByVal Simpan As String) As Boolean
    Dim tmpInfo As String
    Dim FolderBaru As String
    Dim FullFolder As String
    Dim UkPack As String
    Dim FileName As String
    Dim UkHasil As Long
    Dim UkNya As Long
    Dim HndL As Long
    Dim Ext As String
    Dim NamaSama As Boolean
    Dim NamaFull As String
    
    
    YestoAll = False
    NotoAll = False
    RenameAll = False
    Pointer = 0
    PesanError = ""
    

    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        frmUtama.tInfo.Text = "============================================================" & vbCrLf
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Open file " & Alamat & vbCrLf
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "============================================================" & vbCrLf
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Getting File Header" & vbCrLf
        
        If DapatkanHeader(Alamat, hFileR, MySize, Pointer) = False Then
            ExtractArchive = False
            GoTo loncat
        End If

        If Pointer >= MySize Or ErrorHard Then
            ExtractArchive = False
            GoTo loncat
        End If
        
        
start:
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            
            DapatkanInfo hFileR, Pointer, nama
            frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Open => " & nama & vbCrLf
            Attr = InfoJenis.Atribut

            
            '// untuk memastikan format tipe
            If InfoJenis.sFlags And FFile Then
                '// baca Info tentang file
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                Pointer = Pointer + Len(InfoFile)
                

                
                    '// rangkai info
                    With InfoFile
                        CRC = .CRC32
                        UkWal = .UkAwalData
                        UkPack = .UkAhirData
                    End With
                    
                    Attr = InfoJenis.Atribut
                    FileName = TesSlash(Simpan) & nama
                    frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Getting Information => " & nama & vbCrLf
                    If UkWal > 0 Then
                        
                        '// baca data file
                        SetFilePointer hFileR, Pointer, 0, 0
                        ReDim data(InfoFile.UkAhirData - 1)
                        Call ReadFile(hFileR, data(0), InfoFile.UkAhirData, ret, ByVal 0&)
                        
                        If InfoJenis.sFlags And FPassword Then
                            If SetPass = False Then
                                frmPass.Show 1, frmUtama
                                If pass = "" Then GoTo loncat
                            End If
                            '// decript data
                            Decript data, pass
                        End If
                        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Unpacking => " & nama & vbCrLf
                        '// decompress data
                        UkHasil = MyDecompress(data, Hasil, MethodC, UkWal)
                        
                        '// jika nilai crc tidak sama dengan crc awal
                        If CekCheksumK(CRC32(Hasil(), UkWal), CRC, nama) = False Then
                        
                        End If
                        
                    End If
                    
                    '// tulis filenya
                    frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Creating File Information => " & nama & vbCrLf
                    CreateFile FileName, Simpan, UkWal, Hasil, Attr
                    
                Pointer = Pointer + InfoFile.UkAhirData
            Else
                    '// untuk folder
                    NamaFolder = nama
                    FullFolder = Simpan & "\" & NamaFolder
                    frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Creating Directory => " & FullFolder & vbCrLf
                    '// buat foldernya dulu
                    BuatFolderAuto FullFolder
            End If
            
                        
            DoEvents
            If Pointer < MySize Then GoTo start
            
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Finished Extracting " & Alamat & vbCrLf
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "============================================================" & vbCrLf
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "Registration Component" & vbCrLf
        frmUtama.tInfo.Text = frmUtama.tInfo.Text & "============================================================" & vbCrLf
        ExtractArchive = True
loncat:
    Call VbCloseHandle(hFileR)
    
End Function


Private Function Kosongkan(ByRef Info As UntukFile)
    With Info
        .CRC32 = 0
        .CrcPass = 0
        .UkAwalData = 0
        .UkAhirData = 0
    End With
End Function
Private Function RangkaiInfo(CRC As Long, CrcPass As Integer, UkWal As Long, UkHir As Long) As UntukFile
    With InfoFile
        .CRC32 = CRC
        .CrcPass = CrcPass
        .UkAwalData = UkWal
        .UkAhirData = UkHir
    End With
End Function
Public Function RangkaiType(Attr As Integer, TM As Long, Flg As Integer, UkNama As Integer) As InfoType
    With InfoJenis
        .Atribut = Attr
        .Time = TM
        .sFlags = Flg
        .UkNamaFile = UkNama
        .Type_Hash = HashType
    End With
End Function

Private Function UniPath(isi() As Byte, ByVal Flg As Integer) As String
    Dim Path As String
    Dim temp() As Byte
    
    temp = isi
    If Flg And FUnicode Then
        '// untuk tipe unicode
        If CekEncript = True Then
            Decript isi, pass
            isi = StrConv(isi, vbUnicode)
            UniPath = StrConv(isi, vbFromUnicode)
            Encript isi, pass
        Else
            isi = StrConv(isi, vbUnicode)
            UniPath = StrConv(isi, vbFromUnicode)
        End If
        
        UniPath = Left$(UniPath, Len(UniPath))
    Else
        '// untuk tipe Ansi
        '// rubah ke betuk string
        Path = StrConv(isi, vbUnicode)
        
        If CekEncript = True Then
            Decript isi, pass
            Path = StrConv(isi, vbUnicode)
            Encript isi, pass
        End If

        UniPath = StripNulls(Path)
    End If
    
    isi = temp
End Function
Public Function BagianAkhir(Ptr As Long, ByVal HndR As Long, ByVal HndW As Long, BT() As Byte) As Long
    
    'Dapatkan Data File
    SetFilePointer HndR, Ptr, 0, 0
    
    If InfoFile.UkAhirData > 0 Then
        ReDim data(InfoFile.UkAhirData - 1)
        Call ReadFile(HndR, data(0), InfoFile.UkAhirData, ret, ByVal 0&)
    Else
    
    End If
    
    'Tulis TypeNya
    Call WriteFile(HndW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
    'Tulis Nama File
    Call VbWriteFileB(HndW, InfoJenis.UkNamaFile, BT)
    'Tulis Informasi File
    Call WriteFile(HndW, InfoFile, Len(InfoFile), ret, ByVal 0&)
    ' Tulis Data File
    Call VbWriteFileB(HndW, InfoFile.UkAhirData, data)
                    
    End Function
Public Function TulisPesan(ByVal Hnd As Long, Ptr As Long) As Long
    Call WriteFile(Hnd, InfoCommand, Len(InfoCommand), ret, ByVal 0&)
    Ptr = Ptr + Len(InfoCommand)
    Call WriteFile(Hnd, isi(0), InfoCommand.LenPack, ret, ByVal 0&)
    Ptr = Ptr + InfoCommand.LenPack
End Function
Public Function DapatkanInfo(ByVal Hnd As Long, Ptr As Long, nama As String) As Boolean
    Dim MyHash As Integer
    
    DapatkanInfo = True
    'If Khusus = False Then
        '// dapatkan informasi file
        SetFilePointer Hnd, Ptr, 0, 0
        Call ReadFile(Hnd, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
        Ptr = Ptr + Len(InfoJenis)
    
        MyHash = HashType
        
        
        On Error GoTo Exx
        '// dapatkan nama
        ReDim isi(InfoJenis.UkNamaFile - 1)
        SetFilePointer Hnd, Ptr, 0, 0
        Call ReadFile(Hnd, isi(0), InfoJenis.UkNamaFile, ret, ByVal 0&)
        Ptr = Ptr + InfoJenis.UkNamaFile
        '// cocokkan Hash_header
    'End If
                                    
    '// periksa dulu format nama
    
    nama = UniPath(isi, InfoJenis.sFlags)
    If InfoJenis.Type_Hash <> MyHash Or MyHash = 0 Then
        PesanError = PesanError & StripNulls(Alamat) & " Header file " & nama & " telah rusak atau dimodifikasi" & vbCrLf
        DapatkanInfo = False
            'GoTo ExX
    End If

Back:

Exit Function
Exx:
    DapatkanInfo = False
    GoTo Back
End Function
Public Function DapatkanHeader(ByVal Path As String, Hnd As Long, Size As Long, Ptr As Long) As Boolean
    Dim UkHasil As Long
    Dim penanda As String
    Dim Bit() As Byte
    Dim data As String
    Dim ps As Long
    penanda = Chr(&H3C) & Chr(&H61) & Chr(&H35) & Chr(&H63) & Chr(&H99) & Chr(&H3) & Chr(&H64) & Chr(&H33)

'<a5c™d3
    Hnd = CreateFileW(StrPtr(Path), &H80000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
        If Hnd = -1 Then
            DapatkanHeader = False
            GoTo Exx
        End If
        '// dapatkan ukuran file archive
        Size = VbFileLen(Hnd)
        
        '// pisahkan dulu
        ReDim Bit(Size)
        data = Space$(Size)
        SetFilePointer Hnd, 0, 0, 0
        Call VbReadFileB(Hnd, 1, Size, Bit)
        data = StrConv(Bit, vbUnicode)

        ps = InStr(data, penanda)
        If (ps > 0) Then

        Else
            DapatkanHeader = False
            GoTo Exx:
        End If
        
        Ptr = ps + Len(penanda) - 1
        
        '// dapatkan header archive dulu
        SetFilePointer Hnd, Ptr, 0, 0
        Call ReadFile(Hnd, Header, Len(Header), ret, ByVal 0&)

        Ptr = Ptr + Len(Header)
        
        MethodC = Header.MethodCompress
        
        DapatkanHeader = True
        
        '// cocokkan struktur headernya
        If CekHeader = False Then
            'DapatkanHeader = False
            GoTo Exx:
        End If
        
        '// Untuk Command
        If Header.flags And fCommand Then
        
            '// Dapatkan InfoCommand
            SetFilePointer Hnd, Ptr, 0, 0
            Call ReadFile(Hnd, InfoCommand, Len(InfoCommand), ret, ByVal 0&)
            Ptr = Ptr + Len(InfoCommand)
            
            '// Baca Command
            ReDim isi(InfoCommand.LenPack - 1)
            SetFilePointer Hnd, Ptr, 0, 0
            Call ReadFile(Hnd, isi(0), InfoCommand.LenPack, ret, ByVal 0&)
            Ptr = Ptr + InfoCommand.LenPack
            
            '// decompress Command
            UkHasil = MyDecompress(isi, Hasil, MethodC, CLng(InfoCommand.LenUnPack))
            pesan = StrConv(Hasil, vbUnicode)

        Else
            '
        End If
        
        
        '// Jika Archive di encript
        If CekEncript = True Then
            
            '// cocokkan password
            If CocokkanPass = False Then
                DapatkanHeader = False
            End If
        End If
Exx:
End Function
Public Function CekHeader() As Boolean
    Dim MyHash As Integer
    
    ErrorHard = False
    CekHeader = True
    
    
    
    MyHash = HashHeader                      ' HASH HEADER
        
    '// cocokkan signature
        If Header.Signature <> DAFASignature Then
            PesanError = PesanError & StripNulls(Alamat) & " Signature header tidak cocok, kemungkinan bukan merupakan format archive !" & vbCrLf
            ErrorHard = True
            CekHeader = False
            GoTo Exx
        End If
        
        '// cocokkan Hash_header
        If Header.Hash_Head <> MyHash Then
            PesanError = PesanError & StripNulls(Alamat) & " Header archive telah rusak atau dimodifikasi " & vbCrLf
            CekHeader = False
        End If
        
        '// cocokkan format
        If Header.StructurArchive > VersiArchive Then
            PesanError = PesanError & StripNulls(Alamat) & " Merupakan format archive terbaru, silahkan update software ini !" & vbCrLf
            ErrorHard = True
            CekHeader = False
            GoTo Exx
        End If
        
        '// cocokkan metode compress
        If MethodC > VersiCompres Then
            PesanError = PesanError & StripNulls(Alamat) & " Dipack dengan metode terbaru, silahkan update software ini !" & vbCrLf
            CekHeader = False
        End If
        
        '// cocokkan solid archive
        If Header.flags And FSolid Then
            PesanError = PesanError & StripNulls(Alamat) & " Dipack dengan teknik solid archive, untuk versi ini belum mendukung !" & vbCrLf
            CekHeader = False
        End If
Exx:

End Function

Public Sub ShowInformation()
        On Error Resume Next
        With iArc
            Rat = Left(CStr(Round(TotPack / TotUk * 100, 2)), 5)
            .Versi = "0." & Str$(Header.StructurArchive)
            .TotFile = Str$(JumFile)
            .TotPack = Str$(TotPack) & " Bytes"
            .TotSize = Str$(TotUk) & " Bytes"
            .Ratio = Str$(Rat) & " %"
            .Rat = Rat
            .SFX = "0 Bytes"
            .Dictionary = "0 Bytes"
            .Recovary = "0 Bytes"
            .Verification = "Absent"
            
            If Header.flags And FPassword Then
                .Password = "Present"
            Else
                .Password = "Absent"
            End If
            
            If Header.flags And fCommand Then
                .Command = "Present"
            Else
                .Command = "Absent"
            End If
            
            If Header.flags And FLock Then
                .Lock = "Present"
            Else
                .Lock = "Absent"
            End If
            
        End With
End Sub
Public Function CekCheksum(ByVal Nilai1 As Long, ByVal Nilai2 As Long) As Boolean

    '// jika checksum tidak cocok
    If Nilai1 <> Nilai2 Then
        PesanError = PesanError & StripNulls(Alamat) & " Nilai CRC file tidak cocok kemungkinan terjadi karena Password salah !" & vbNewLine
        CekCheksum = False
        
    Else
        CekCheksum = True
    End If
End Function
Public Function CekCheksumK(ByVal Nilai1 As Long, ByVal Nilai2 As Long, ByVal nama As String) As Boolean

    '// jika checksum tidak cocok
    If Nilai1 <> Nilai2 Then
        PesanError = PesanError & Hex$(Nilai1) & " <> " & Hex$(Nilai2) & " " & StripNulls(Alamat) & " : File " & nama & " Nilai crc salah file rusak !" & vbNewLine
        CekCheksumK = False
    Else
        CekCheksumK = True
    End If
End Function
Public Function CocokkanPass() As Boolean
    isi = StrConv(pass, vbFromUnicode)
    '// jika checksum tidak cocok
    If CekCheksum(Header.CrcEncript, GetCrc16(isi, UBound(isi))) = False Then
        CocokkanPass = False
    Else
        CocokkanPass = True
    End If
End Function
Public Function MyCompress(Bit() As Byte, Bit2() As Byte, ByVal Method As Byte, ByVal Lev As Long) As Long
    Select Case Method
        Case MetfLz
            MyCompress = Press(Bit, Bit2)
    End Select
End Function
Public Function MyDecompress(Bit() As Byte, Bit2() As Byte, ByVal Method As Byte, ByVal Uk As Long) As Long
    Select Case Method
        Case MetfLz
            MyDecompress = Dec(Bit, Bit2, Uk)
    End Select
    
End Function
Public Function CekEncript() As Boolean
'// jika file di encript
    If Header.flags And FEncript Then
        CekEncript = True
    Else
        CekEncript = False
    End If
End Function
Public Function CekPassword() As Boolean
'// jika file di encript
    If InfoJenis.sFlags And FPassword Then
        CekPassword = True
    Else
        CekPassword = False
    End If
End Function
Public Function Falidkah(ByVal nama As String, ByVal FullPath As String) As Boolean
    Dim Pos1 As Long
    Dim Pos2 As Long
    Dim Pot As String
    
    Pos1 = InStr(nama, FullPath)
    Pos2 = InStr(nama, "\")
    
    If nama = FullPath Then
        Falidkah = True
    ElseIf Pos1 = 1 Then
        Pot = Mid$(nama, Len(FullPath) + 1, 1)
        If Pot = "\" Then
            Falidkah = True
        Else
            Falidkah = False
        End If
    Else
        Falidkah = False
    End If
End Function
Public Function GetData(ByVal Path As String, ByRef Size As Long) As String
    Dim Tmp() As Byte
    Dim Hnd As Long
    Dim ret As Long
    
    Hnd = CreateFileW(StrPtr(Path), &H80000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
        Size = VbFileLen(Hnd)
        
        ReDim Tmp(Size - 1)
        SetFilePointer Hnd, 0, 0, 0
        Call ReadFile(Hnd, Tmp(0), Size, ret, ByVal 0&)
    CloseHandle Hnd
    
    GetData = StrConv(Tmp, vbUnicode)
End Function
Public Function CreateFile(ByVal Path As String, ByVal Simpan As String, ByVal SizeAwal As Long, DataNya() As Byte, ByVal Attr As Integer) As Long
    Dim HndL As Long
    Dim Ext As String
    
            '// Jika File sudah ada
            If PathFileExists(StrPtr(Path)) = 1 Then
            
                '// Jika belum diberi perintah
                If YestoAll = False And NotoAll = False And RenameAll = False Then
                
                    '// dapatkan ukuran file dahulu
                    HndL = VbOpenFile(Path)
                        UkNya = VbFileLen(HndL)
                    VbCloseHandle HndL
                    
                    With frmReplace
                        '// tulis informasinya
                        .lblnama = nama
                        .lUk1.Caption = UkNya & " bytes"
                        .lUk2.Caption = InfoFile.UkAwalData & " bytes"
                            
                        '// gambah iconya
                        Ex_Icon Path, .Pic1, ico64
                        Ex_Icon Path, .Pic2, ico64
                        NameRename = Path
                            
                        ' Tampilkan
                        .Show 1
                        Path = NameRename
                    End With
                        
                '// jika perintah rename
                ElseIf RenameAll = True Then
                    Ext = AmbilExtensi(nama)
                    Path = Left$(nama, Len(nama) - Len(Ext)) & "(1)" & Ext
                    Path = TesSlash(Simpan) & Path
                End If
            End If
            
            If NotoAll = False And Path <> "" Then
                BuatFolderAuto AmbilAlamat(Path)
                    
                '/ buka dulu
                hFileW = CreateFileW(StrPtr(Path), &H40000000, &H2, ByVal 0&, 2, 0, 0)
                    '// jika ada datanya
                    If SizeAwal > 0 Then
                        Call VbWriteFileB(hFileW, SizeAwal, DataNya)
                    End If
                '// tutup file
                Call VbCloseHandle(hFileW)
                Call SetFileAttr(Path, Attr)
            End If

End Function
Private Sub CleanCons()
    JumFile = 0
    JumDir = 0
    JumObject = 0
    Flg = 0
    Offset = 0
    
    Pointer = 0
    TotPack = 0
    TotUk = 0
    PesanError = ""
    

End Sub
Public Function EditLock(ByVal AlamatNya As String) As Boolean
    
        If DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) = False Then
            EditLock = False
        Else
            Header.flags = Header.flags Or FLock
            Header.Hash_Head = HashHeader

        End If
        
        CloseHandle hFileR
        
        hFileW = CreateFileW(StrPtr(AlamatNya), &H40000000, &H2, ByVal 0&, 4, 0, 0)
            SetFilePointer hFileW, 0, 0, 0
            Call WriteFile(hFileW, Header, Len(Header), ret, ByVal 0&)
        CloseHandle hFileW

End Function

