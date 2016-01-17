Attribute VB_Name = "ModDesain"
Dim Alamat As String

Private Enum eBitmapResources
    bmpLarge256 = 101
End Enum
Public Function Desain(lokasi As FrmUtama)
    Dim mm As cBand
    With lokasi.t
        With .Item(0)
            Set .ImageList = NewImageList(48, 38, imlColor32)
            
            For i = 0 To 19
                .ImageList.AddFromDc lokasi.pic(i).hDC, 48, 38
            Next i
            
            .Buttons.Add "Add", "Add", , 0
            .Buttons.Add "Extract to", "Extract to", , 1
            .Buttons.Add "Test", "Test", , 2
            .Buttons.Add "View", "View", , 3
            .Buttons.Add "Delete", "Delete", , 4
            .Buttons.Add "Find", "Find", , 5
            .Buttons.Add "Print", "Print", , 6
            .Buttons.Add "Wizard", "Wizard", , 7
            .Buttons.Add "Convert", "Convert", , 8
            .Buttons.Add "Info", "Info", , 9
            .Buttons.Add "Exit", "Exit", , 10
            .Buttons.Add "Repair", "Repair", , 11
            .Buttons.Add "Extract", "Extract", , 12
            .Buttons.Add "VirusScan", "VirusScan", , 13
            .Buttons.Add "Comment", "Comment", , 14
            .Buttons.Add "Protect", "Protect", , 15
            .Buttons.Add "Lock", "Lock", , 16
            .Buttons.Add "SFX", "SFX", , 17
            .Buttons.Add "Report", "Report", , 18
            .Buttons.Add "New", "New", , 19

       End With
       With .Item(1)
            Set .ImageList = NewImageList(18, 18, imlColor32)
            .ImageList.AddFromDc lokasi.PicRe.hDC, 18, 18
            .Buttons.Add , "", , 0
       End With
    End With
    
    With FrmFindR.t
        With .Item(0)
            Set .ImageList = NewImageList(48, 38, imlColor32)
            For i = 0 To 6
                .ImageList.AddFromDc lokasi.pic(i).hDC, 48, 38
            Next i
            .Buttons.Add , "Extract to", , 2
            .Buttons.Add , "View", , 4
            .Buttons.Add , "Locate", , 6

       End With
    End With
    
        With lokasi.LVSelect
            .Font.source = fntSourceSysMenu
            .Columns.Add 1, "Lokasi", , , lvwAlignLeft, 5000
            .Columns.Add 2, "Posisi", , , lvwAlignLeft, 5000
            .Columns.Add 3, "Posisi", , , lvwAlignLeft, 5000
        End With
        
        With lokasi.LVRead
            .Font.source = fntSourceSysMenu
            Set .ImageList(lvwImageSmallIcon) = NewImageList(16, 16, imlColor32)
            Set .ImageList(lvwImageLargeIcon) = NewImageList(32, 32, imlColor32)
            .Columns.Add 1, "Nama", , , lvwAlignLeft, 2000
            .Columns.Add 2, "Alamat", , , lvwAlignLeft, 3000
            .Columns.Add 3, "Ukuran Awal", , , lvwAlignLeft, 2000
            .Columns.Add 4, "Ukuran Pack", , , lvwAlignLeft, 2000
            .Columns.Add 5, "Ratio Pack", , , lvwAlignLeft, 2000
            .Columns.Add 6, "Type", , , lvwAlignLeft, 3000
            .Columns.Add 7, "Atribut", , , lvwAlignLeft, 1000
            .Columns.Add 8, "Nilai Crc32", , , lvwAlignLeft, 2000
            .Columns.Add 9, "", , , lvwAlignLeft, 0
        End With
        
        With FrmFindR.lvFind
            .Font.source = fntSourceSysMenu
            Set .ImageList = NewImageList(16, 16, imlColor32)
            .Columns.Add 1, "File", , , lvwAlignLeft, 3500
            .Columns.Add 2, "Location", , , lvwAlignLeft, 1500
            .Columns.Add 3, "Context", , , lvwAlignLeft, 5000
            .Columns.Add 4, "", , , lvwAlignLeft, 0
        End With
        
        With lokasi.TV
            Call .Initialize
            Call .InitializeImageList
            Call .AddIcon(lokasi.Pic2.Picture)
            Call .AddIcon(lokasi.pic3.Picture)
            Call .AddIcon(lokasi.pic4.Picture)
            .HasButtons = True
            .HasLines = True
            .HasRootLines = True
            .TrackSelect = True
            .BackColor = vbWhite
            .ForeColor = &H0
            .LineColor = vbBlack
        End With

        With lokasi.status
            Set .ImageList = gImageListSmall
            
            .Panels.Add , , , sbarStandard, , 11, 700, 1000
            .Panels.Add , , , sbarStandard, , , 2000, 1000
            .Panels.Add , , , sbarStandard, , , 4000, 1000
            .Panels.Add , , , sbarStandard, , , 5000, 1000
        End With
        
    lokasi.r.Bands.Add lokasi.t(0), , , True, True
    lokasi.r.Bands.Add lokasi.t(1), , , , True
    'lokasi.r.Bands.Add lokasi.TAlamat, , , True
    
    FrmFindR.r.Bands.Add FrmFindR.t(0)
    Call DisableTool
End Function
Public Sub DisableTool()
    With FrmUtama
        .SaveCopy.Enabled = False
        .Copyfils.Enabled = False
        .Past.Enabled = False
        .SelGrop.Enabled = False
        .DesGrop.Enabled = False
        .InvSel.Enabled = False
        .Wiz.Enabled = False
        .ScanArc.Enabled = False
        .ConvArc.Enabled = False
        .RepArc.Enabled = False
        .ConvToSFX.Enabled = False
        .GeneratRep.Enabled = False
        .Petunjuk.Enabled = False
        '.BukaNotepad.Enabled = False
        .PrintFF.Enabled = False
        .RepArca.Enabled = False
        .ImEx.Enabled = False
        .Home.Enabled = False
        .FltFolderV.Enabled = False
        .ViewLg.Enabled = False
        .ClearL.Enabled = False
        '.CreteFol.Enabled = False
        .ProtArcFromD.Enabled = False
        .LvX.Enabled = False
        .DetailX.Enabled = False
        '.Favor.Enabled = False
    End With
    
End Sub
