Attribute VB_Name = "ModMain"
Sub Main()
    Call BuatFolderTemp
    Call DaftarReg
    Call InitCrc16
    Call InitCrc
    Call InitPack

    Call Desain(FrmUtama)
    
    Call LoadSetting
    Call CreateSetting
    Call Mulai
    
End Sub
