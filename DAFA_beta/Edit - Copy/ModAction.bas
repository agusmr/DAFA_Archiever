Attribute VB_Name = "ModAction"
Public OptFokus As Integer
Public Sub TambahkanFile()
    Dim Fokus As String
    
    If Alamat <> "" Then
        With FrmUtama.c
            .FileFilter = "Semua File|*.*"
            If .ShowOpen Then
                If .FileName <> "" Then
                    FrmUtama.LVSelect.ListItems.Clear
                    Set G = FrmUtama.LVSelect.ListItems.Add(, StripNulls(.FileName))
                        G.SubItem(2).Text = 1
                    TambahArchive Alamat, 1
                End If
            End If
        End With
    Else
        MsgBox "Tidak ada file archive yang terbuka !", vbCritical
    End If
End Sub

Public Sub TambahkanFolder()
    Dim Fokus As String
    Dim JumSelect As Long
    
    If Alamat <> "" Then
        FrmBrow.Show 1
            If tSimpan <> "" Then
                FrmUtama.LVSelect.ListItems.Clear
                Kumpulkan TesSlash(tSimpan), -1, AmbilAlamat(tSimpan)
                JumSelect = FrmUtama.LVSelect.ListItems.Count
                TambahArchive Alamat, JumSelect
            End If
    Else
        MsgBox "Tidak ada file archive yang terbuka !", vbCritical
    End If
End Sub

Public Sub ExtractSpesial()
    Dim JumSelect As Long
    
    tSimpan = TesSlash(AmbilAlamat(Alamat))
    FrmBrow.Show 1
    If tSimpan <> "" Then
        JumSelect = MasukSelect
        If JumSelect > 0 Then
            ExtractFile Alamat, JumSelect, tSimpan
        End If
    End If
End Sub
Public Sub BacaArsip()
    Dim nama As String
    
    FrmUtama.LVSelect.ListItems.Clear
    nama = CariSelect
    ViewArc nama
End Sub
Public Sub HapusArsip()
    Dim JumSelect As Long
    Dim Fokus As String
    
    JumSelect = MasukSelect
    If JumSelect > 0 Then
        If MsgBox("Apakah yakin ingin hapus " & JumSelect & " object ?", vbYesNo) = vbYes Then
            HapusArchive Alamat, JumSelect
        End If
    Else
        MsgBox "Pilih file atau folder yang ingin dihapus !", vbCritical
    End If
End Sub
Public Sub RenameArsip()
    Dim nama As String
    Dim Fokus As String
    
    nama = CariSelect
    If nama <> "" Then
        RenameArchive Alamat, nama
    End If
End Sub
Public Sub SetPassword()
    frmPass.Show 1, FrmUtama
    If pass <> "" Then
        SetPass = True
        FrmUtama.status.Panels(1).Text = "+"
    Else
        SetPass = False
        FrmUtama.status.Panels(1).Text = "-"
    End If
End Sub
Public Sub ExtractAllSpesial()
    Dim JumSelect As Long
    
    
    If Alamat <> "" Then
        JumSelect = MasukSelect
        tSimpan = TesSlash(AmbilAlamat(Alamat))
        FrmBrow.Show 1
        If tSimpan <> "" Then
            If JumFile < 1 Then GetArchive Alamat, JumDir, JumFile
            
            If JumSelect > 0 Then
                ExtractSingleFol Alamat, CariSelect, tSimpan
            Else
                ExtractArchive Alamat, tSimpan, 0, True
            End If
        End If
    Else
        MsgBox "Tidak ada file archive yang terbuka !", vbCritical
    End If
End Sub
Public Sub CariFile()
    If Alamat <> "" Then
        FrmFind.Show
    Else
        MsgBox "Tidak ada file archive yang terbuka !", vbCritical
    End If
End Sub
Public Sub LoadFolder()
    Dim hNode As Long
    Dim NamaKey As String
    
    
    NamaKey = TesSlash(CariSelect)
    If ApakahFile = True Then
        Call BacaArsip
    Else
        hNode = FrmUtama.TV.GetKeyNode(NamaKey)
        'MasukkanTree Alamat, hNode
        Tampilkan NamaKey
    End If
End Sub
Public Sub KeluarArsip()
    Alamat = ""
    FrmUtama.tAlamat.Text = ""
    FrmUtama.LVRead.ListItems.Clear
    FrmUtama.TV.Clear
    EnableButton False
    Terbuka = False
End Sub
Public Sub ExtractAllArsip()
    Dim JumSelect As Long
    
    If Alamat <> "" Then
        JumSelect = MasukSelect
        tSimpan = TesSlash(AmbilAlamat(Alamat))
        If tSimpan <> "" Then
            If JumSelect > 0 Then
                ExtractArchive Alamat, tSimpan, JumSelect
            Else
                ExtractArchive Alamat, tSimpan, 0, True
            End If
        End If
    Else
        MsgBox "Tidak ada file archive yang terbuka !", vbCritical
    End If
End Sub
Public Sub PilihSemua()
    With FrmUtama.LVRead.ListItems
        For i = 1 To .Count
            .Item(i).Selected = True
        Next i
    End With
End Sub
Public Sub BukaArsip()
    With FrmUtama.c
        .FileFilter = "Darma Archive|*.gus"
        If .ShowOpen Then
            Alamat = StripNulls(.FileName)
            FrmUtama.tAlamat.Text = Alamat
            BacaArchive Alamat
        End If
    End With
End Sub
Public Sub ShowInfo()
    ' Info =1
    ' Opt = 2
    ' Comm =3
    ' sfx = 4
    ' 5 = OPT + Lock
    OptFokus = 1
    frmOption.Show 1
End Sub
Public Sub ShowProtect()
    OptFokus = 2
    frmOption.Show 1
End Sub
Public Sub ShowLock()
    OptFokus = 5
    frmOption.Show 1
End Sub
Public Sub ShowComment()
    OptFokus = 3
    frmOption.Show 1
    
End Sub
Public Function GantiView(ByVal Tipe As eListViewStyle)
    'MasukkanTree Alamat, FokusNode
    FrmUtama.LVRead.View = Tipe
End Function
Public Sub UntukLock()
    With FrmUtama
        With .t.Item(0)
            .Buttons(1).Enabled = False
            .Buttons(5).Enabled = False
            '.Buttons(c).Enabled = False
            
        End With
        .TambahFile.Enabled = False
        .AddFol.Enabled = False
        .UbahNama.Enabled = False
        .HapusFile.Enabled = False
        .CreteFol.Enabled = False
        .AddArcCom.Enabled = False
    End With
End Sub
Function tampilkanOpt(Index As Integer) As Long

    With frmOption
        
        Select Case Index
            
            Case 1
                .fInfo.Visible = True
            Case 2
                .fInfo.Visible = False
                .fOption.Visible = True
            Case 3
                .fInfo.Visible = False
                .fOption.Visible = False
                .fComment.Visible = True
                .tComment.Text = Pesan
                .tComment.SelStart = 0
                .tComment.SelLength = Len(Pesan)
                'tComment.SetFocus
            Case 4
                .fInfo.Visible = False
                .fOption.Visible = False
                .fComment.Visible = False
                .fSFX.Visible = True
        End Select
    End With
End Function
Public Sub TestArchives()
    If TestArchive(Alamat) = True Then MsgBox "No error found during test operation !"
End Sub
