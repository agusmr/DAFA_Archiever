Attribute VB_Name = "ModInit"
Public Sub Daftarkan()
    Call BuatFolderTemp
    Call DaftarReg
    Call BuildTableCrc32
    Call Desain(FrmUtama)
End Sub
