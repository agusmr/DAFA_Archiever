Attribute VB_Name = "ModPacked"
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

'===========================================================================
'===========================================================================
Private Type CommandArchive
    CommandCrc      As Integer      ' BERISI CRC KOMENTAR
    LenUnPack       As Integer      ' BERISI UKURAN KOMENTAR SEBELUM DI PACK
    LenPack         As Integer      ' BERISI UKURAN KOMENTAR SETELAH DI PACK
End Type
'===========================================================================

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
Public Pesan        As String
Public MethodC      As Byte
Public RatioC       As Long
Public hFileDibuat As Long
Public ErrorHard As Boolean

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

Public Function BuatArchive(ByVal Alamat As String, _
                            ByVal Simpan As String, _
                            ByVal status As Label, _
                            Optional satu As Boolean = False) As Boolean
                            
    ' // PENJELASAN VALUE //////////////////////////////
    ' // Alamat : LOKASI FILE / FOLDER YANG INGIN DI BUNDEL
    ' // Simpan : LOKASI TEMPAT PENYIMPANAN ARCHIVE
    ' // Status : HANYA UNTUK UPDATE STATUS AJA
    ' // Satu   : JIKA BERNIALAI TRUE MAKA MENANDAKAN HANYA AKAN MEMBUNDEL SATU FILE
    '           : JIKA BERNIALAI FALSE MAKA MENANDAKAN AKAN MEMBUNDEL FILE DALAM FOLDER
    
    
    '// BERSIHKAN DULU KONSTANTA DAN COMPONENT LAINYA
    Call CleanCons
    BuatArchive = True
    
    With frmProses
        .Show
        .lblWaktu.Caption = ": 00:00:00"
    End With
    
    'FrmUtama.TAlamat.text = Simpan
    FrmUtama.Tim1.Enabled = True

    tSimpan = Simpan        ' <=== ini simpan dulu di variabel tSimpan karena nanti digunakan
    
    status.Caption = "Status         : Menulis Header File"     ' Update Status
    
    Offset = Len(Header)
    
    '// MULAI DISINI SUDAH MEMBUAT FILE ARCHIVE
        '// TULISKAN HEADER ARCHIVE DULU
        If Not BuatHeader(hFileW, Offset, Simpan) Then
            BuatArchive = False
            PesanError = PesanError & "gagal membuat header archive !" & vbCrLf
            GoTo Ex
        End If
        
        hFileDibuat = hFileW            ' simpan untuk jaga jaga bila dibutuhkan
        
        status.Caption = "Status         : Mencari File"            ' Update Status
        
        If satu = False Then
            '// UNTUK MEMBUNDEL BANYAK FILE
            If Right(Alamat, 1) = "\" Then
                Call SearchFile(Alamat)     ' Dapatkan Informasi file yang mau dibundel
                PerbaikiTampilan
            Else
                'frmutama.
            End If
            
            '// KUMPULKAN FILE DAN LANGSUNG DITULIS
            Kumpulkan Alamat, hFileW, AmbilAlamat(Simpan)
        Else
            '// HANYA UNTUK MEMBUNDEL SATU FILE
            If Not TulisFile(hFileW, AmbilNama(Alamat), Alamat, 0, 0) Then
                BuatArchive = False
                PesanError = PesanError & "Gagal menulis file " & StripNulls(Alamat) & " !" & vbCrLf
                GoTo Exx
            End If
        End If
Exx:
    Call VbCloseHandle(hFileW)      ' Tutup File Archive
    
    Alamat = Simpan
    pass = ""
    DiEncrip = 0
    DiLock = 0
    DiPassword = 0
    Pesan = ""
    DiPesan = 0
    MethodC = 0
    BacaArchive Simpan
Ex:
Call ShowErroR(BuatArchive)
End Function
Public Function BuatHeader(Hnd As Long, Offset As Long, ByVal nama As String) As Boolean
    
    BuatHeader = True
    
    '// BUAT DULU FILE ARCHIVE NYA
    Hnd = CreateFileW(StrPtr(nama), &H40000000, &H2, ByVal 0&, 2, 0, 0)
        If Hnd = -1 Then
            BuatHeader = False
            PesanError = PesanError & "Tidak bisa membuat file, mungkin dikunci oleh proses lain !"
            GoTo Ex
        End If
        '// RANGKAI BAGIAN HEADER ARCHIVE
        With Header
            .Signature = DAFASignature                  ' SIGNATURE ARCHIVE
            .StructurArchive = VersiArchive             ' VERSI ARCHIVE
            '// UNTUK JAGA JAGA KALAU INFORMASI METODE ARCHIVE KOSONG
            If MethodC = 0 Then
                RatioC = 9
                MethodC = MetDeflate
            End If
            
            .MethodCompress = MethodC                   ' METHODE ARCHIVE
            
            '// INI UNTUK SCURITY
            If DiPassword And DiEncrip Then
                Flg = Flg Or FEncript                   ' ADD FLAGS ENCRIPT
                isi = StrConv(pass, vbFromUnicode)      ' RUBAH PASSWORD KEBENTUK BYTE
                .CrcEncript = GetCrc16(isi, UBound(isi))  ' CHECKSUM NAMA PASSWORD
            Else
                .CrcEncript = 0                         ' JIKA TANPA DI ENCRIPT
            End If
            
            If DiLock Then                              ' JIKA ARCHIVE DI KUNCI
                Flg = Flg Or FLock                      ' ADD FLAGS LOCK
            End If
            
            If RatioC <> 0 Then
                Flg = Flg Or FCompress
            End If
            
            '// RANGKAI KHUSUS UNTUK BAGIAN COMMENT
            If DiPesan Then
                Flg = Flg Or fCommand                                   ' ADD FLAGS COMMENT
                isi = StrConv(Pesan, vbFromUnicode)
                With InfoCommand
                    .CommandCrc = GetCrc16(isi, UBound(isi))            ' CHEKSUM DATA COMMENT
                    .LenUnPack = Len(Pesan)                             ' UKURAN DATA COMMENT AWAL
                    .LenPack = MyCompress(isi, Hasil, MethodC, RatioC)  ' UKURAN DATA COMMENT SETELAH DI COMPRESS
                End With
            End If
            .flags = Flg                                                ' FLAGS
            .Hash_Head = HashHeader
        End With
        
        Call WriteFile(Hnd, Header, Offset, ret, ByVal 0&)  ' TULIS BAGIAN HEADER ARHIVE
        
        '// NAH INI UNTUK PENAMBAHAN COMMENT
        If DiPesan = True Then
            Call WriteFile(Hnd, InfoCommand, Len(InfoCommand), ret, ByVal 0&)   ' TULIS INFO COMMENT
            Offset = Offset + Len(InfoCommand)
            
            Call WriteFile(Hnd, Hasil(0), InfoCommand.LenPack, ret, ByVal 0&)   ' TULIS DATA COMMENT
            Offset = Offset + InfoCommand.LenPack
        End If
Ex:
    Call ShowErroR(BuatHeader)
End Function
Public Function TulisFile(ByVal hFileW As Long, ByVal PathSimpan As String, ByVal NamaNya As String, ByVal TM As Long, ByVal Attr As Integer) As Boolean
    Dim NamaPot As String
    Dim UkHasil As Long
    Dim CRC As Long
    Dim UkNama As Integer
    Dim Dir As String
    Dim hFileR As Long
    
    Flg = 0
    TulisFile = True
    
    '// BUKA DULU FILE YANG AKAN DI BUNDEL
    hFileR = CreateFileW(StrPtr(NamaNya), &H80000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
        If hFileR = -1 And ValidFolder(NamaNya) = False Then
            TulisFile = False
            PesanError = PesanError & "Gagal membuka file " & StripNulls(NamaNya) & " !" & vbCrLf
            GoTo Ex
        End If
        
        
        PathSimpan = PerbaikiPath(PathSimpan, Flg)
                
        If ValidFile(NamaNya) And ValidFolder(NamaNya) = False Then
            z = z + 1

            '// UNTUK JENIS FILE
            
            Flg = Flg Or FFile                                      ' ADD FLAGS FILE
            
            '// UNTUK FLASH SAJA
            With frmProses
                .TAlamat.Text = AmbilNama(PathSimpan)
                .lblUkuran.Caption = "Ukuran  : " & ukuRan & " Bytes"
            End With
            
            ukuRan = VbFileLen(hFileR)                              ' DAPATKAN UKURAN FILE
            
            If ukuRan > 0 Then                                      ' UNTUK FILE YANG ADA DATANYA
                
                VbReadFileB hFileR, 1, ukuRan, data                 ' DAPATKAN DATANYA
                
                CRC = CRC32(data(), ukuRan)                        ' HITUNG CRC DATA FILE
                
                UkHasil = MyCompress(data, Hasil, MethodC, 9)  ' COMPRESS DATA
                
                If DiPassword Or CekEncript Then                    ' JIKA ARCHIVE DI KUNCI
                    If Len(pass) > 0 Then
                        Encript Hasil, pass                         ' ENCRIPT DATA
                    Else
                        MsgBox "Password kosong !"
                    End If
                End If
                
                Call RangkaiInfo(CRC, 0, ukuRan, UkHasil)     ' RANGKAI INFORMASI
            Else
                '// UNTUK FILE KOSONG
                Kosongkan InfoFile                                  ' KOSONGKAN INFORMASI
            End If
                                            
            myDoEvents
        ElseIf PathIsDirectory(StrPtr(NamaNya)) Then                   ' UNTUK FOLDER
            Flg = Flg Or FFolder                                    ' ADD FLAGS FOLDER
        End If
                
        UkNama = Len(PathSimpan)                                       ' DAPATKAN UKURAN NAMANYA
        
        isi = StrConv(PathSimpan, vbFromUnicode)
        
        '// DISINI MASIH ADA BUGS UNTUK ENCRIPT UNICODE ////////////////
                
        If DiPassword Or CekEncript Then                            ' JIKA ARCHIVE DI KUNCI
            Flg = Flg Or FPassword                                  ' ADD FLAGS PASSWORD
            If CekEncript Then                                      ' JIKA ARCHIVE DI ENCRIPT
                If Len(pass) > 0 Then
                    Encript isi, pass                               ' ENCRIPT NAMANYA
                Else
                    MsgBox "Password kosong !"
                End If
            End If
        End If
        
        If Flg And FFile Then Attr = GetAttrFile(NamaNya)


        Call RangkaiType(Attr, TM, Flg, UkNama)              ' RANGKAI INFORMASI TYPE
        
        Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
        'Tulis Nama File
        Call VbWriteFileB(hFileW, UkNama, isi)
        '// jika merupakan file
        If Flg And FFile Then
            '// Tulis Informasi File

            InfoFile.OffsetData = Offset
            Call WriteFile(hFileW, InfoFile, Len(InfoFile), ret, ByVal 0&)
            
            'If z = JumFile Then
                'Tulis Data File
                If ukuRan > 0 Then
                    'If z = 1 Then
                    Call WriteFile(hFileW, Hasil(0), UkHasil, ret, ByVal 0&)
                End If
            'End If
            
            Offset = Offset + Len(InfoFile)
            Offset = Offset + UkHasil
            
        End If
        Offset = Offset + Len(InfoJenis) + UkNama
    Call VbCloseHandle(hFileR)
    
Ex:
    Call ShowErroR(TulisFile)
End Function
Public Function BacaArchive(ByVal AlamatNya As String) As Boolean
    Terbuka = True
    BacaArchive = True
    Alamat = AlamatNya  '<== Simpan lagi untuk jaga jaga
    
    ' // kosongkan dulu
    ' ============================================
    CleanCons
    'Pass = ""
    Pesan = ""
    FrmUtama.TAlamat.Text = Alamat
    EnableButton True
    ' ============================================
    On Error GoTo Exx
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If Not DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) Then
            BacaArchive = False
            PesanError = PesanError & "Gagal membuka archive " & StripNulls(AlamatNya) & " !" & vbCrLf
            GoTo Ex
        End If
        
        My = FrmUtama.TV.AddNode(, , , AmbilNama(Alamat), 2, 2)
        
        
        If Pointer >= MySize Or ErrorHard Then
            BacaArchive = False
            GoTo loncat
        End If
    Do
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, NamaFolder
            
            '// untuk memastikan tipe
            If InfoJenis.sFlags And FFolder Then
                JumDir = JumDir + 1
                
                '// Tampilkan Directory
                Masuk NamaFolder
            Else
                JumFile = JumFile + 1
                'Pointer = Pointer + 8
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                
                UkWal = InfoFile.UkAwalData
                UkPack = InfoFile.UkAhirData
                
                TotUk = TotUk + UkWal
                TotPack = TotPack + UkPack
                
                Pointer = Pointer + Len(InfoFile)
                
                Pointer = Pointer + UkPack
            End If
            
            
            'myDoEvents
        Loop While Pointer < MySize
    
loncat:
Exx:
    Call VbCloseHandle(hFileR)
    On Error Resume Next
    'If BacaArchive = True Then
        With FrmUtama
        
            .Show
            .TV.SetFocus
            .LVRead.SetFocus
            .LVRead.SetFocusedItem 1
            With .status
                .Panels.Item(2).Text = "Proses Selesai !"
                .Panels.Item(3).Text = "Jumlah File : " & JumFile & ", " & "Folder : " & JumDir
            End With
        End With
        Call ShowInformation
    'End If
Ex:
    Call ShowErroR(BacaArchive)
End Function
Public Function BacaFile(ByVal AlamatNya As String, ByVal NamaTree As String) As Boolean
    
    Dim NamaFile As String
    Dim NamaFileE As String
    
    BacaFile = True
    FokusNode = NamaTree
    Alamat = AlamatNya

    Pointer = 0
    
    PesanError = ""
    
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If Not DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) Then
            BacaFile = False
            PesanError = PesanError & "Gagal membuka archive " & StripNulls(AlamatNya) & " !" & vbCrLf
            GoTo Ex
        End If
        
        '// jika archive kosong loncati
        If Pointer >= MySize Or ErrorHard Then
            BacaFile = False
            GoTo loncat
        End If
        
        On Error GoTo loncat
                    
Do
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            Attr = InfoJenis.Atribut
            
            '// untuk memastikan tipe
            If InfoJenis.sFlags And FFolder Then
                '// jika tipe folder
                
                If InStr(nama, NamaTree) > 0 Then
                    '// untuk file didalam folder
                    Masukkan nama, 0, 0, Attr, 0, Folder, 0

                End If
            Else
                '// untuk tipe file
                                
                NamaFile = NamaTree & AmbilNama(nama)
                                
                '// dapatkan informasi file tambahan
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                
                '// jika nama sesuai dengan folder yang dipilih
                If nama = NamaFile Then
                    '// rangkai informasi
                    With InfoFile
                        CRC = .CRC32
                        UkWal = .UkAwalData
                        UkPack = .UkAhirData
                        Offset = .OffsetData
                    End With
                    
                    '// masukkan ke list view
                    Masukkan nama, UkWal, UkPack, Attr, CRC, File, Offset
                End If
                
                Pointer = Pointer + Len(InfoFile)
                Pointer = Pointer + InfoFile.UkAhirData
            End If
            
            'myDoEvents
            'If Pointer < MySize Then GoTo start
            Loop While Pointer < MySize
loncat:
    Call VbCloseHandle(hFileR)
Ex:
    Call ShowErroR(BacaFile)
End Function
Public Function FindArchive(ByVal AlamatNya As String, ByVal FileFind As String) As Boolean
    
    Dim NamaFile As String
    Dim NamaFileE As String
    
    FokusNode = NamaTree
    Alamat = AlamatNya
    FrmUtama.TAlamat.Text = Alamat
    FindArchive = True
    Pointer = 0
    PesanError = ""
    
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If Not DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) Then
            FindArchive = False
            
            GoTo loncat
        End If
        
        If Pointer >= MySize Or ErrorHard Then
            FindArchive = False
            GoTo loncat
        End If
start:

            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            
            '// untuk memastikan tipe
            If InfoJenis.sFlags And FFolder Then
            
                '// jika folder yang dicari ketemu
                If InStr(nama, FileFind) > 0 Then
                    '// masukkan ketabel list view
                    MasukFind nama, AlamatNya, "", 0
                End If
                
            Else
                '// jika merupakan file
                NamaFile = NamaTree & AmbilNama(nama)
                
                '// dapatkan ukuran akhir data
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                
                '// jika file yang dicari ketemu
                If InStr(NamaFile, FileFind) > 0 Then
                    '// masukkan ketabel list view
                    MasukFind nama, AlamatNya, "", InfoFile.OffsetData
                End If
                
                Pointer = Pointer + Len(InfoFile)
                Pointer = Pointer + InfoFile.UkAhirData
            End If
            
            'myDoEvents
            If Pointer < MySize Then GoTo start
        FindArchive = True
loncat:
    Call VbCloseHandle(hFileR)
    
    If FindArchive = True Then
        FrmFindR.Show
    End If
End Function
Public Function GetArchive(ByVal AlamatNya As String, Folder As Long, File As Long) As Boolean
    
    Dim NamaFile As String
    Dim NamaFileE As String
    
    
    Alamat = AlamatNya
    
    Pointer = 0
    PesanError = ""
    JumDir = 0
    JumFile = 0
    
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) = False Then
            GetArchive = False
            GoTo loncat
        End If
        
                If Pointer >= MySize Or ErrorHard Then
            GetArchive = False
            GoTo loncat
        End If
start:

            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            
            '// untuk memastikan tipe
            If InfoJenis.sFlags And FFolder Then
                JumDir = JumDir + 1
                
            Else
                JumFile = JumFile + 1
                
                '// dapatkan ukuran akhir data
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                
                Pointer = Pointer + Len(InfoFile)
                Pointer = Pointer + InfoFile.UkAhirData
            End If
            
            'myDoEvents
            If Pointer < MySize Then GoTo start
        GetArchive = True
loncat:
    Call VbCloseHandle(hFileR)
    
    Folder = JumDir
    File = JumFile
    
End Function
Public Function TestArchive(ByVal AlamatNya As String) As Boolean
    
    Dim NamaFile As String
    Dim NamaFileE As String
    Dim UkHasil As Long
    
    Alamat = AlamatNya
    
    Pointer = 0
    PesanError = ""
    TestArchive = True
    frmProses.Show
    
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) = False Then
            TestArchive = False
            GoTo loncat
        End If
        
                If Pointer >= MySize Or ErrorHard Then
            TestArchive = False
            GoTo loncat
        End If
start:

            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            
            '// untuk memastikan tipe
            If InfoJenis.sFlags And FFolder Then
                '
            Else
            
                With frmProses
                    .TAlamat.Text = AmbilNama(nama)
                    .lblUkuran.Caption = "Ukuran  : " & UkWal & " Bytes"
                End With
                
                '// dapatkan ukuran akhir data
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                Pointer = Pointer + Len(InfoFile)
                
                '// rangkai info
                With InfoFile
                    CRC = .CRC32
                    UkWal = .UkAwalData
                    UkPack = .UkAhirData
                End With
                
                '// untuk jaga jaga aja
                If UkWal > 0 Then
                
                    'Dapatkan Data File
                    SetFilePointer hFileR, Pointer, 0, 0
                    ReDim data(InfoFile.UkAhirData - 1)
                    Call ReadFile(hFileR, data(0), UkPack, ret, ByVal 0&)
                    
                    '// jika terkunci
                    If InfoJenis.sFlags And FPassword Then
                        If SetPass = False Then
                            frmPass.Show 1, FrmUtama
                            SetPass = True
                            If pass = "" Then GoTo loncat
                        End If
                        '// decript data
                        Decript data, pass
                    End If
                    
                    '// decompress data
                    UkHasil = MyDecompress(data, Hasil, MethodC, UkWal)
                    
                    '// cocokkan checksum
                    If CekCheksumK(CRC32(Hasil(), UkWal), CRC, nama) = False Then
                        TestArchive = False
                    End If
                        
                Else
                    '// kosongkan aja
                    ReDim Hasil(0)
                    Hasil(0) = 0
                    UkHasil = 0
                End If
                
                Pointer = Pointer + InfoFile.UkAhirData
            End If
            
            'myDoEvents
            If Pointer < MySize Then GoTo start
loncat:
    Call VbCloseHandle(hFileR)
    
    Unload frmProses
End Function

Public Function ExtractArchive(ByVal Alamat As String, ByVal Simpan As String, ByVal JumSelect As Long, Optional Semua As Boolean = False) As Boolean
    
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
    
    With frmProses
        .Show
        .lblStatus = "Status         : Menulis File"
    End With
    
    YestoAll = False
    NotoAll = False
    RenameAll = False
    Pointer = 0
    PesanError = ""
            
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
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
            
            Attr = InfoJenis.Atribut
            ' untuk multi select
            If Semua = False Then
                For x = 1 To JumSelect
                    NamaFull = PotongSlash(FrmUtama.LVSelect.ListItems(x).Text)
                    If Falidkah(nama, NamaFull) = True Then
                        NamaSama = True
                        Exit For
                    Else
                        NamaSama = False
                    End If
                Next x
            Else
                NamaSama = True
            End If
            
            '// untuk memastikan format tipe
            If InfoJenis.sFlags And FFile Then
                '// baca Info tentang file
                SetFilePointer hFileR, Pointer, 0, 0
                Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                Pointer = Pointer + Len(InfoFile)
                
                '// tampilkan
                frmProses.TAlamat.Text = nama

                If NamaSama = True Then
                    '// rangkai info
                    With InfoFile
                        CRC = .CRC32
                        UkWal = .UkAwalData
                        UkPack = .UkAhirData
                    End With
                    
                    Attr = InfoJenis.Atribut
                    FileName = TesSlash(Simpan) & nama
                    
                    If UkWal > 0 Then
                        
                        '// baca data file
                        SetFilePointer hFileR, Pointer, 0, 0
                        ReDim data(InfoFile.UkAhirData - 1)
                        Call ReadFile(hFileR, data(0), InfoFile.UkAhirData, ret, ByVal 0&)
                        
                        If InfoJenis.sFlags And FPassword Then
                            If SetPass = False Then
                                frmPass.Show 1, FrmUtama
                                If pass = "" Then GoTo loncat
                            End If
                            '// decript data
                            Decript data, pass
                        End If
                        '// decompress data
                        UkHasil = MyDecompress(data, Hasil, MethodC, UkWal)
                        
                        '// jika nilai crc tidak sama dengan crc awal
                        If CekCheksumK(CRC32(Hasil(), UkWal), CRC, nama) = False Then
                        
                        End If
                        
                    End If
                    
                    '// tulis filenya
                    
                    CreateFile FileName, Simpan, UkWal, Hasil, Attr
                End If
                Pointer = Pointer + InfoFile.UkAhirData
            Else
                If NamaSama = True Then
                    '// untuk folder
                    NamaFolder = nama
                    FullFolder = Simpan & "\" & NamaFolder
                    '// buat foldernya dulu
                    BuatFolderAuto FullFolder
                End If
            End If
            
            With frmProses
                .TAlamat.Text = AmbilNama(nama)
                .lblUkuran.Caption = "Ukuran  : " & UkWal & " Bytes"
            End With
                        
            'myDoEvents
            If Pointer < MySize Then GoTo start
            
        ExtractArchive = True
loncat:
    Call VbCloseHandle(hFileR)
    
    Unload frmProses
End Function
Public Function ExtractSingleFol(ByVal AlamatNya As String, ByVal NameFolder As String, ByVal Simpan As String) As Long
    
    Dim Tmp As Long
    Dim pos As Long
    Dim PosMem As Long
    Dim PosDisk As Long
    Dim PosDat As Long
    Dim BT() As Byte
    Dim v() As Byte
    Dim DAT As String
    Dim Uk As Long
    Dim Path As String
    Dim PathFix As String
    Dim hFileW As Long
    
    
    DAT = GetData(AlamatNya, Uk)
    
    YestoAll = False
    NotoAll = False
    RenameAll = False
    PesanError = ""

    Pointer = 1
    Tmp = 0
    BT = StrConv(DAT, vbFromUnicode)
    PosDat = VarPtr(BT(0))
    With frmProses
        .Show
        .lblStatus = "Status         : Menulis File"
    End With
    Do
        pos = InStr(Pointer, DAT, NameFolder)
        If pos > 0 Then
            PosDisk = pos - Len(InfoJenis)
            PosMem = PosDat + PosDisk - 1
            
            Call CopyMem(ByVal VarPtr(InfoJenis), ByVal PosMem, Len(InfoJenis))
            PosMem = PosMem + Len(InfoJenis)
            
            nama = Space$(InfoJenis.UkNamaFile)
            Call CopyMem(ByVal StrPtr(nama), ByVal PosMem, InfoJenis.UkNamaFile)
            nama = StrConv(nama, vbUnicode)
            frmProses.TAlamat.Text = nama
            If InfoJenis.sFlags And FFile Then

                PosMem = PosMem + InfoJenis.UkNamaFile
                Call CopyMem(ByVal VarPtr(InfoFile), ByVal PosMem, Len(InfoFile))
                
                PosMem = PosMem + Len(InfoFile)
                                
                PathFix = Replace$(nama, FokusNode, "")
                Path = TesSlash(Simpan) & PathFix
                
                BuatFolderAuto AmbilAlamat(Path)

                UkWal = InfoFile.UkAwalData
                Attr = InfoJenis.Atribut
                If UkWal > 0 Then
                    ReDim v(InfoFile.UkAhirData - 1)
                    Call CopyMem(ByVal VarPtr(v(0)), ByVal PosMem, InfoFile.UkAhirData)
                        
                    MyDecompress v, Hasil, 1, UkWal
                End If
                
                CreateFile Path, Simpan, UkWal, Hasil, Attr
                
                Pointer = pos + InfoJenis.UkNamaFile + Len(InfoFile) + InfoFile.UkAhirData
            Else
                Pointer = pos + InfoJenis.UkNamaFile
                    '// untuk folder
                    NamaFolder = nama
                    FullFolder = Simpan & "\" & NamaFolder
                    '// buat foldernya dulu
                    BuatFolderAuto FullFolder
            End If

                        
            Tmp = Tmp + 1
        Else
            GoTo Keluar
        End If
            With frmProses
                .TAlamat.Text = AmbilNama(nama)
                .lblUkuran.Caption = "Ukuran  : " & UkWal & " Bytes"
            End With
    Loop While Pointer < Uk
Keluar:
Unload frmProses
End Function
Public Function ExtractFile(ByVal Alamat As String, ByVal JumSelect As Long, ByVal Simpan As String) As Boolean
    
    Dim FolderBaru As String
    Dim UkPack As String
    Dim FileName As String
    Dim NamaFull As String
    Dim NamaFullE As String
    Dim NamaSama As Boolean
    Dim UkHasil As Long
    
    Pointer = 0
    PesanError = ""
    
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(Alamat, hFileR, MySize, Pointer) = False Then
            ExtractFile = False
            GoTo loncat
        End If
        
        '// jika archive kosong loncati
                If Pointer >= MySize Or ErrorHard Then
            ExtractFile = False
            GoTo loncat
        End If
        '// untuk multi select
        For x = 1 To JumSelect
            NamaFull = FrmUtama.LVSelect.ListItems(x).Text
            Pointer = CLng(FrmUtama.LVSelect.ListItems(x).SubItem(2).Text)
                    
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            '// untuk memastikan tipe file
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
                    
                '// baca data file
                SetFilePointer hFileR, Pointer, 0, 0
                    
                '// untuk jaga jaga aja
                If UkWal <> 0 Then
                    ReDim data(InfoFile.UkAhirData - 1)
                    Call ReadFile(hFileR, data(0), InfoFile.UkAhirData, ret, ByVal 0&)
                    '// jika terkunci
                    If InfoJenis.sFlags And FPassword Then
                        If SetPass = False Then
                            frmPass.Show 1, FrmUtama
                            If pass = "" Then GoTo loncat
                        End If
                        '// decript data
                        Decript data, pass
                    End If
                    
                    '// decompress data
                    UkHasil = MyDecompress(data, Hasil, MethodC, UkWal)
                        
                    '// cocokkan checksum
                    If CekCheksumK(CRC32(Hasil(), UkWal), CRC, nama) = False Then
                        ExtractFile = False
                        GoTo loncat
                    End If
                        
                Else
                    '// kosongkan aja
                    ReDim Hasil(0)
                    Hasil(0) = 0
                    UkHasil = 0
                End If
                    
                '// Tulis File
                hFileW = CreateFileW(StrPtr(Simpan & AmbilNama(nama)), &H40000000, &H2, ByVal 0&, 1, 0, 0)
                    If UkWal > 0 Then
                        Call VbWriteFileB(hFileW, UkWal, Hasil)
                    End If
                Call VbCloseHandle(hFileW)
                        
            End If
            
        Next x
        ExtractFile = True
loncat:
    Call VbCloseHandle(hFileR)
End Function
Public Function HapusArchive(ByVal Alamat As String, ByVal JumSelect As Long) As Boolean
    
    Dim NamaFull As String
    Dim NamaSama As Boolean
    Dim hFileR As Long
    Dim x As Long
'On Error Resume Next

    Pointer = 0
    PesanError = ""
    Offset = 0
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(Alamat, hFileR, MySize, Pointer) = False Then
            HapusArchive = False
            GoTo loncat
        End If
        
        With frmProses
            .Show
            .lblStatus.Caption = "Proses   : Copy File !"
        End With
        
        '// buat file temporary
        hFileW = CreateFileW(StrPtr(Alamat & ".tmp"), &H40000000, &H2, ByVal 0&, 1, 0, 0)
            '// Mulai Tulis Header
            Call WriteFile(hFileW, Header, Len(Header), ret, ByVal 0&)
            Offset = Len(Header)
            
            '// tulis Pesan
            If Header.flags And fCommand Then
                Call TulisPesan(hFileW, Offset)
            End If
            
            '// jika archive kosong loncati
                    If Pointer >= MySize Or ErrorHard Then
                HapusArchive = False
                GoTo loncat
            End If
            
start:
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
                
                ' untuk multi select
                    For x = 1 To JumSelect
                        NamaFull = PotongSlash(FrmUtama.LVSelect.ListItems(x).Text)
                        If Falidkah(nama, NamaFull) = True Then
                            NamaSama = True
                            Exit For
                        Else
                            NamaSama = False
                        End If
                    Next x
                    
                '// untuk memastikan tipe file
                If InfoJenis.sFlags And FFolder Then
                    If NamaSama = False Then
                        Offset = Offset + Len(InfoJenis) + InfoJenis.UkNamaFile
                        'Tulis TypeNya
                        Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
                        'Tulis Nama folder
                        Call VbWriteFileB(hFileW, InfoJenis.UkNamaFile, isi)
                    End If
                    
                Else
                '// tampilkan
                frmProses.TAlamat.Text = nama
                '// baca Info tentang file
                    SetFilePointer hFileR, Pointer, 0, 0
                    Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                    Pointer = Pointer + Len(InfoFile)
                    
                    If NamaSama = False Then
                        InfoFile.OffsetData = Offset
                        BagianAkhir Pointer, hFileR, hFileW, isi
                        Offset = Offset + Len(InfoJenis) + _
                                 Len(InfoFile) + InfoJenis.UkNamaFile + InfoFile.UkAhirData
                    End If
                    
                    Pointer = Pointer + InfoFile.UkAhirData
                    
                End If
            
        If Pointer < MySize Then GoTo start
        
        HapusArchive = True
loncat:
            Call VbCloseHandle(hFileW)
        Call VbCloseHandle(hFileR)
        
    If HapusArchive = True Then
        Hapus Alamat
        RenameFile StrPtr(Alamat & ".tmp"), StrPtr(Alamat)
        myDoEvents
        frmProses.lblStatus.Caption = "Info  : Hapus file !"
    Else
        Hapus StrPtr(Alamat & ".tmp")
    End If
    
    Fuokus = FokusNode
    Unload frmProses
    BacaArchive Alamat
    TampilkanTV Fuokus
End Function

Public Function RenameArchive(ByVal Alamat As String, ByVal NamaFull As String) As Boolean
    
    Dim Folder As Boolean
    Dim FolderLawas As String
    Dim FolderNew As String
    Dim RenName As String
    Dim NamaSama As Boolean
    
    Pointer = 0
    PesanError = ""
    Offset = 0
    Folder = False
    
    RenName = InputBox("Masukkan Nama Baru", "Rename File Archive", AmbilNama(NamaFull))
    
    If RenName = "" Or rename = AmbilNama(NamaFull) Then
        RenameArchive = False
        GoTo loncat
    End If

    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(Alamat, hFileR, MySize, Pointer) = False Then
            RenameArchive = False
            GoTo loncat
        End If
        
        '// buat file Temporary
        hFileW = CreateFileW(StrPtr(Alamat & ".tmp"), &H40000000, &H2, ByVal 0&, 1, 0, 0)
            ' Mulai Tulis Header
            Call WriteFile(hFileW, Header, Len(Header), ret, ByVal 0&)
            Offset = Len(Header)
            
            '// tulis Pesan
            If Header.flags And fCommand Then
                Call TulisPesan(hFileW, Offset)
            End If
                    
                        
        '// jika archive kosong loncati
        If Pointer >= MySize Or ErrorHard Then
            RenameArchive = False
            GoTo loncat
        End If
        
start:
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            ' untuk multi select
                If Falidkah(nama, NamaFull) = True Then
                    NamaSama = True
                Else
                    NamaSama = False
                End If
            
                If InfoJenis.sFlags And FFile Then
                    If NamaSama = True Then
                        If Folder = False Then
                            If nama <> "" Then
                                If FokusNode = "" Then
                                    nama = RenName
                                Else
                                    nama = AmbilAlamat(NamaFull) & RenName
                                End If
                            Else
                                nama = NamaFull
                            End If
                        Else
                            nama = FolderNew & Mid$(nama, Len(FolderLawas) + 1)
                        End If
                        
                        If InfoJenis.sFlags And FUnicode Then
                            isi = StrConv(nama, vbUnicode)
                            isi = StrConv(isi, vbFromUnicode)
                        Else
                            isi = StrConv(nama, vbFromUnicode)
                        End If
                        '// jika file di encript
                        If Header.flags And FEncript Then _
                            Encript isi, pass
                            
                        InfoJenis.UkNamaFile = UBound(isi) + 1
                        
                    End If
                    
                    '// baca Info tentang file
                    SetFilePointer hFileR, Pointer, 0, 0
                    
                    Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                    Pointer = Pointer + Len(InfoFile)
                    InfoJenis.Type_Hash = HashType
                    
                    InfoFile.OffsetData = Offset
                    
                    
                    BagianAkhir Pointer, hFileR, hFileW, isi

                    Pointer = Pointer + InfoFile.UkAhirData
                    Offset = Offset + Len(InfoJenis) + _
                            Len(InfoFile) + InfoJenis.UkNamaFile + InfoFile.UkAhirData
                Else
                    If NamaSama = True Then
                        
                        If Folder = False Then
                            Folder = True
                            FolderLawas = nama
                            
                            If nama <> "" Then
                                If FokusNode = "" Then
                                    nama = RenName
                                Else
                                    nama = AmbilAlamat(NamaFull) & RenName
                                End If
                            Else
                                nama = NamaFull
                            End If
                            FolderNew = nama
                            nama = FolderNew
                        Else
                            nama = FolderNew & Mid$(nama, Len(FolderLawas) + 1)
                        End If
                        
                        If InfoJenis.sFlags And FUnicode Then
                            isi = StrConv(nama, vbUnicode)
                            isi = StrConv(isi, vbFromUnicode)
                        Else
                            isi = StrConv(nama, vbFromUnicode)
                        End If
                        '// jika file di encript
                        If Header.flags And FEncript Then _
                            Encript isi, pass
                            
                        InfoJenis.UkNamaFile = UBound(isi) + 1
                    End If
                    
                    'Tulis TypeNya
                    InfoJenis.Type_Hash = HashType
                    Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
                    'Tulis Nama folder
                    Call VbWriteFileB(hFileW, InfoJenis.UkNamaFile, isi)
                    
                    Offset = Offset + Len(InfoJenis) + InfoJenis.UkNamaFile
                End If
                
        If Pointer < MySize Then GoTo start
        RenameArchive = True
loncat:
        Call VbCloseHandle(hFileW)
    Call VbCloseHandle(hFileR)
    
    If RenameArchive = True Then
        Hapus Alamat
        RenameFile StrPtr(Alamat & ".tmp"), StrPtr(Alamat)
        myDoEvents
        frmProses.lblStatus.Caption = "Info  : Hapus file !"
        Unload frmProses
        BacaArchive Alamat
        Fuokus = FokusNode
        TampilkanTV Fuokus

    End If
    
    
    
End Function
Public Function TambahArchive(ByVal Alamat As String, ByVal jumlah As Long) As Boolean
    
    Dim NamaFull As String
    Dim NamaSama As Boolean
    Dim Sama As Long
    Dim i&, x&, z&
    Dim UkHasil As Long
    Dim MyCrc As Long
    Dim Dir As String
    Dim G As cListItem
    Dim NamaNya As String

    '// kosongkan dulu
    Pointer = 0
    Sama = 0
    PesanError = ""
    Offset = 0
    
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(Alamat, hFileR, MySize, Pointer) = False Then
            TambahArchive = False
            GoTo loncat
        End If
        
        If Header.flags And FLock Then
            Call VbCloseHandle(hFileR)
            PesanError = "File Archive tidak bisa dimodifikasi karena terkunci !"
            TambahArchive = False
            Call ShowErroR(TambahArchive)
            GoTo Ex
        End If
        
        '// masukkan status
        With frmProses
            .Show
            .lblStatus.Caption = "Proses   : Copy File !"
        End With
        
        
        '// Buat file temporery
        hFileW = CreateFileW(StrPtr(Alamat & ".tmp"), &H40000000, &H2, ByVal 0&, 1, 0, 0)
            ' Mulai Tulis Header
            Call WriteFile(hFileW, Header, Len(Header), ret, ByVal 0&)
            Offset = Len(Header)
            
            '// tulis Pesan
            If Header.flags And fCommand Then
                Call TulisPesan(hFileW, Offset)
            End If
            
            If Pointer >= MySize Or ErrorHard Then
                'TambahArchive = False
                GoTo loncat
            End If
        
start:
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            
            For x = 1 To jumlah
                With FrmUtama.LVSelect.ListItems(x)
                    NamaNya = .SubItem(3).Text
                    
                    NamaFull = FokusNode & NamaNya
                    'MsgBox nama & "|" & NamaFull
                    If nama = NamaFull Then
                        MsgBox "nama sama"
                        .SubItem(2).Text = 0
                        NamaFull = .Text
                        NamaSama = True
                        Exit For
                    Else
                        NamaSama = False
                    End If
                End With
            Next x
                    
                If InfoJenis.sFlags And FFile Then
                    frmProses.TAlamat.Text = nama
                    
                    '// baca Info tentang file
                    SetFilePointer hFileR, Pointer, 0, 0
                    Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                    Pointer = Pointer + Len(InfoFile)
                    
                    If Not NamaSama Then
                        InfoFile.OffsetData = Offset
                        BagianAkhir Pointer, hFileR, hFileW, isi
                        Offset = Offset + Len(InfoJenis) + _
                        Len(InfoFile) + InfoJenis.UkNamaFile + InfoFile.UkAhirData
                    Else
                        TulisFile hFileW, nama, NamaFull, 0, 0
                    End If
                    
                    Pointer = Pointer + InfoFile.UkAhirData
                Else
                    'Tulis TypeNya
                    Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
                    'Tulis Nama folder
                    Call VbWriteFileB(hFileW, InfoJenis.UkNamaFile, isi)
                    Offset = Offset + Len(InfoJenis) + InfoJenis.UkNamaFile
                End If
                
        If Pointer < MySize Then GoTo start
loncat:
    Call VbCloseHandle(hFileR)
    
    TambahArchive = True
    
    
    ' Buka File Baru
    If jumlah > 1 Then
        JumFile = jumlah
    End If
    
    For i = 1 To jumlah
        Flg = 0
        With FrmUtama.LVSelect.ListItems(i)
            nama = .Text
            'frmProses.lblStatus.Caption = "Proses  : Tambah File"
            '// buka file yang ingin ditambah
            If CInt(.SubItem(2).Text) = 1 Then
                If jumlah = 1 Then
                    TulisFile hFileW, FokusNode & AmbilNama(nama), nama, 0, 0
                Else
                    TulisFile hFileW, FokusNode & .SubItem(3).Text, nama, 0, 0
                End If
            End If
        End With
    Next i
    
Call VbCloseHandle(hFileW)
    If TambahArchive = True Then
        Call Hapus(Alamat)
        RenameFile StrPtr(Alamat & ".tmp"), StrPtr(Alamat)
    End If
Ex:
    
    Unload frmProses
    Fuokus = FokusNode
    Unload frmProses
    BacaArchive Alamat
    TampilkanTV Fuokus

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
Private Function PerbaikiPath(ByVal Path As String, ByRef Flg As Integer) As String
    Dim x$
    
    x = UniToAnsi(Path)
    
    If InStr(x, "?") > 0 Then
        '// berarti tipe unicode
        Flg = Flg Or FUnicode
        PerbaikiPath = StrConv(Path, vbUnicode)
    Else
        '// berarti tipe ansi
        PerbaikiPath = Path
    End If
End Function
Private Function UniPath(isi() As Byte, ByVal Flg As Integer) As String
    Dim Path As String
    Dim Temp() As Byte
    
    Temp = isi
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
    
    isi = Temp
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
    Call ShowErroR(DapatkanInfo)
Exit Function
Exx:
    DapatkanInfo = False
    GoTo Back
End Function
Public Function DapatkanHeader(ByVal Path As String, Hnd As Long, Size As Long, Ptr As Long) As Boolean
    Dim UkHasil As Long
    
    Hnd = CreateFileW(StrPtr(Path), &H80000000, &H1 Or &H2, ByVal 0&, 3, 0, 0)
        If Hnd = -1 Then
            DapatkanHeader = False
            GoTo Exx
        End If
        '// dapatkan ukuran file archive
        Size = VbFileLen(Hnd)
        
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
            Pesan = StrConv(Hasil, vbUnicode)

        Else
            '
        End If
        
        '// jika archive dikunci
        If Header.flags And FLock Then
            Call UntukLock
        End If
        
        '// Jika Archive di encript
        If CekEncript = True Then
            If pass = "" Then
                frmPass.Show 1, FrmUtama

                If pass <> "" Then
                    isi = StrConv(pass, vbFromUnicode)
                    SetPass = True
                    FrmUtama.status.Panels(1).Text = "+"
                Else
                    SetPass = False
                    FrmUtama.status.Panels(1).Text = "-"
                End If
            End If
            
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

        Call ShowErroR(CekHeader)
End Function
Public Sub ShowErroR(ByVal Valid As Boolean)
    If Not Valid Then
        FrmDiagnosa.TPesan.Text = PesanError
        FrmDiagnosa.Show
    End If
End Sub
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
        Call ShowErroR(CekCheksum)
    Else
        CekCheksum = True
    End If
End Function
Public Function CekCheksumK(ByVal Nilai1 As Long, ByVal Nilai2 As Long, ByVal nama As String) As Boolean

    '// jika checksum tidak cocok
    If Nilai1 <> Nilai2 Then
        PesanError = PesanError & Hex$(Nilai1) & " <> " & Hex$(Nilai2) & " " & StripNulls(Alamat) & " : File " & nama & " Nilai crc salah file rusak !" & vbNewLine
        CekCheksumK = False
        Call ShowErroR(CekCheksumK)
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
        Case MetDeflate
            MyCompress = CompressByte(Bit, Bit2, Lev)
        Case MetBzib2
            MyCompress = z2CompressData(Bit, Bit2, Lev)
        Case MetLZMA
            Call LZMACompress_Simple(Bit, Bit2, MyCompress, Lev)
        Case MetAPlib
            MyCompress = CompressByte1(Bit, Bit2)
        Case MetfLz
            MyCompress = Press(Bit, Bit2)
    End Select
End Function
Public Function MyDecompress(Bit() As Byte, Bit2() As Byte, ByVal Method As Byte, ByVal Uk As Long) As Long
    Select Case Method
        Case MetDeflate
            MyDecompress = DecompressByteArray(Bit, Bit2, Uk)
        Case MetBzib2
            MyDecompress = z2DeCompressData(Bit, Bit2, Uk)
        Case MetLZMA
            Call LZMADecompress_Simple(Bit, Bit2, Uk)
        Case MetAPlib
            MyDecompress = DecompressByte1(Bit, Bit2, Uk)
        Case MetfLz
            MyDecompress = DecompressByte2(Bit, Bit2, Uk)
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
    FrmUtama.LVRead.ListItems.Clear
    FrmUtama.TV.Clear
    PesanError = ""
    
    FrmUtama.Tim1.Enabled = False
    Unload frmProses
    Unload FrmPilih

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

Public Function EditCommand(ByVal AlamatNya As String, ByVal Teks As String) As Boolean
    Dim TempC() As Byte
    Dim sTempC As Long
    
    Pointer = 0
    If DapatkanHeader(AlamatNya, hFileR, MySize, Pointer) = False Then
        EditCommand = False
    Else
        If Teks = "" Then
            Header.flags = Header.flags Xor fCommand
        Else
            Header.flags = Header.flags Or fCommand
        End If
    End If
    Header.Hash_Head = HashHeader
        '// buat file temporary
        hFileW = CreateFileW(StrPtr(AlamatNya & ".tmp"), &H40000000, &H2, ByVal 0&, 1, 0, 0)
            '// Mulai Tulis Header
            Call WriteFile(hFileW, Header, Len(Header), ret, ByVal 0&)
            
            TempC = StrConv(Teks, vbFromUnicode)
            sTempC = UBound(TempC)
            
            Offset = Len(Header)
            '// tulis Pesan
            If Header.flags And fCommand Then
                With InfoCommand
                    .CommandCrc = GetCrc16(TempC, sTempC)
                    .LenUnPack = sTempC + 1
                    .LenPack = MyCompress(TempC, isi, 1, 9)
                End With
                Call TulisPesan(hFileW, Offset)
            End If
            
            '// jika archive kosong loncati
            If Pointer >= MySize Or ErrorHard Then
                EditCommand = False
                MsgBox Pointer & " " & MySize
                GoTo loncat
            End If
            
start:
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
                
                    
                '// untuk memastikan tipe file
                If InfoJenis.sFlags And FFolder Then
                    Offset = Offset + Len(InfoJenis) + InfoJenis.UkNamaFile
                    'Tulis TypeNya
                    Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
                    'Tulis Nama folder
                    Call VbWriteFileB(hFileW, InfoJenis.UkNamaFile, isi)
                    
                Else
                    '// tampilkan
                    frmProses.TAlamat.Text = nama
                    
                    '// baca Info tentang file
                    SetFilePointer hFileR, Pointer, 0, 0
                    Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                    Pointer = Pointer + Len(InfoFile)
                    
                    InfoFile.OffsetData = Offset
                    BagianAkhir Pointer, hFileR, hFileW, isi
                    Offset = Offset + Len(InfoJenis) + _
                                 Len(InfoFile) + InfoJenis.UkNamaFile + InfoFile.UkAhirData
                    
                    Pointer = Pointer + InfoFile.UkAhirData
                    
                End If
            
        If Pointer < MySize Then GoTo start
        
        EditCommand = True
loncat:
            Call VbCloseHandle(hFileW)
        Call VbCloseHandle(hFileR)
        
    If EditCommand = True Then
        Hapus Alamat
        RenameFile StrPtr(Alamat & ".tmp"), StrPtr(Alamat)
        myDoEvents
        frmProses.lblStatus.Caption = "Info  : Hapus file !"
    Else
        Hapus StrPtr(Alamat & ".tmp")
    End If
    
    Fuokus = FokusNode
    Unload frmProses
    BacaArchive Alamat
    Call LoadPesan
    TampilkanTV Fuokus
End Function
Public Function MakeNewFolder(ByVal Alamat As String) As Boolean

    Dim Folder As Boolean
    Dim FolderLawas As String
    Dim FolderNew As String
    Dim NewFolder As String
    Dim NamaSama As Boolean
    
    Pointer = 0
    PesanError = ""
    Offset = 0
    Folder = False
    
    NewFolder = InputBox("Masukkan Nama Folder Baru", "Create new folder", "New Folder")
    
    If NewFolder = "" Then
        MakeNewFolder = False
        GoTo loncat
    End If
    
    NewFolder = FokusNode & NewFolder
    '// buka file archive dulu
        '// dapatkan MySize
                    'Header
                    'Pointer
        If DapatkanHeader(Alamat, hFileR, MySize, Pointer) = False Then
            MakeNewFolder = False
            GoTo loncat
        End If
        
        '// buat file Temporary
        hFileW = CreateFileW(StrPtr(Alamat & ".tmp"), &H40000000, &H2, ByVal 0&, 1, 0, 0)
            ' Mulai Tulis Header
            Call WriteFile(hFileW, Header, Len(Header), ret, ByVal 0&)
            Offset = Len(Header)
            
            '// tulis Pesan
            If Header.flags And fCommand Then
                Call TulisPesan(hFileW, Offset)
            End If
                    
                        
            '// jika archive kosong loncati
                    If Pointer >= MySize Or ErrorHard Then
                MakeNewFolder = False
                GoTo loncat
            End If
        
start:
            '// Dapatkan informasi awal
            '// Info Jenis
            '// nama file
            DapatkanInfo hFileR, Pointer, nama
            ' untuk multi select
                If Falidkah(nama, NewFolder) = True Then
                    NamaSama = True
                Else
                    NamaSama = False
                End If
            
                '// untuk memastikan tipe file
                If InfoJenis.sFlags And FFolder Then
                    Offset = Offset + Len(InfoJenis) + InfoJenis.UkNamaFile
                    'Tulis TypeNya
                    Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
                    'Tulis Nama folder
                    Call VbWriteFileB(hFileW, InfoJenis.UkNamaFile, isi)
                    
                Else
                    '// tampilkan
                    frmProses.TAlamat.Text = nama
                    
                    '// baca Info tentang file
                    SetFilePointer hFileR, Pointer, 0, 0
                    Call ReadFile(hFileR, InfoFile, Len(InfoFile), ret, ByVal 0&)
                    Pointer = Pointer + Len(InfoFile)
                    
                    InfoFile.OffsetData = Offset
                    BagianAkhir Pointer, hFileR, hFileW, isi
                    Offset = Offset + Len(InfoJenis) + _
                                 Len(InfoFile) + InfoJenis.UkNamaFile + InfoFile.UkAhirData
                    
                    Pointer = Pointer + InfoFile.UkAhirData
                    
                End If
        If Pointer < MySize Then GoTo start
        
        If NamaSama = False Then

            isi = StrConv(NewFolder, vbFromUnicode)
            With InfoJenis
                .Atribut = vbNormal
                Attr = FFolder
                If Header.flags And FEncript Then
                    Attr = Attr Or FEncript
                    Encript isi, pass
                End If
                .sFlags = Attr
                .Time = 0
                .UkNamaFile = UBound(isi) + 1
            End With
            
            'Tulis TypeNya
            InfoJenis.Type_Hash = HashType
            Call WriteFile(hFileW, InfoJenis, Len(InfoJenis), ret, ByVal 0&)
            'Tulis Nama folder
            Call VbWriteFileB(hFileW, InfoJenis.UkNamaFile, isi)
        End If
        
        MakeNewFolder = True
loncat:
        Call VbCloseHandle(hFileW)
    Call VbCloseHandle(hFileR)
    
    If MakeNewFolder = True Then
        Hapus Alamat
        RenameFile StrPtr(Alamat & ".tmp"), StrPtr(Alamat)
        myDoEvents
        frmProses.lblStatus.Caption = "Info  : Hapus file !"
        Unload frmProses
        BacaArchive Alamat
        Fuokus = FokusNode
        TampilkanTV Fuokus

    End If
    
End Function


