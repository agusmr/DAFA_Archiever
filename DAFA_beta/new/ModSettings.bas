Attribute VB_Name = "ModSettings"
Dim data As String
Dim Control() As String
Dim cnt() As String
Dim Lawas As String

Const Head = "{DAFA CONTROLER}"

Public Sub LoadSetting()
    Dim i As Long
    Dim ShowMain As Boolean
    Dim ButtonUp As Boolean
    Dim ShowText As Boolean
    Dim ShowAddress As Boolean
    Dim LargeButton As Boolean
    Dim LockToolbar As Boolean
    Dim Sama As Boolean
    Dim x As Long
    Dim nama As String
    Dim temp() As String
    
    ShowMain = False
    ButtonUp = False
    ShowAddress = False
    ShowText = False
    LargeButton = False
    LockToolbar = False
    Sama = False
    
    Open App.Path & "\Settings.txt" For Binary As #1
        data = Space$(LOF(1))
        Get #1, , data
    Close #1
    
    Control = Split(data, vbNewLine)
    With FrmUtama.LVRead
    
        cnt = Split(Control(5), "=")
        FrmSetting.cLog.Value = CByte(cnt(1))
        
        cnt = Split(Control(7), "=")
        SetView CByte(cnt(1))
        .View = CByte(cnt(1))
        
        cnt = Split(Control(8), "=")
        FrmSetting.cShowGrid.Value = CByte(cnt(1))
        .GridLines = CByte(cnt(1))
        
        cnt = Split(Control(9), "=")
        FrmSetting.cFullselect.Value = CByte(cnt(1))
        .FullRowSelect = CByte(cnt(1))
        
        '// Toolbar
        With FrmSetting
            For i = 0 To 5
                cnt = Split(Control(i + 11), "=")
                .cToolbar(i).Value = CByte(cnt(1))
                'If cnt(1) = 1 Then
                    Select Case i
                        Case 0
                            ShowMain = True
                             FrmUtama.r.Bands(1).Visible = CBool(cnt(1))
                        Case 1
                            '
                        Case 2
                            FrmUtama.tAlamat.Visible = CBool(cnt(1))
                        Case 3
                            '
                        Case 4
                            '
                        Case 5
                            FrmUtama.r.Bands(1).Gripper = Not CBool(cnt(1))
                            FrmUtama.r.Bands(2).Gripper = Not CBool(cnt(1))
                    End Select
                'End If
            Next i
        End With
        
        With FrmUtama.t(0)
            For i = 1 To 19
                cnt = Split(Control(i + 17), "=")
                .Buttons(i).Visible = CByte(cnt(1))
                FrmSetting.cButton(i - 1).Value = CByte(cnt(1))
            Next i
        End With
        
        FrmSetting.tTemp.Text = Control(38)
        
        If InStr(Control(40), "|") Then
            Lawas = Control(40)
            cnt = Split(Control(40), "|")
            For i = 0 To UBound(cnt) - 1
                For x = 1 To FrmUtama.MnFavo.count - 1
                    If FrmUtama.MnFavo(x).Tag = cnt(i) Then
                        Sama = True
                        Exit For
                    End If
                Next x
                If Sama = False Then
                    FrmUtama.tAlamat.AddItem cnt(i)
                    AddFavorit cnt(i), "", 2
                End If
            Next i
        End If
        
        If InStr(Control(42), "*") Then
            TemPFavo = Control(42)
            cnt = Split(Control(42), "*")
            For i = 0 To UBound(cnt) - 1
                For x = 1 To FrmUtama.mnuFavorites.count - 1
                    If FrmUtama.mnuFavorites(x).Tag = cnt(i) Then
                        Sama = True
                        Exit For
                    End If
                Next x
                If Sama = False Then
                    temp = Split(cnt(i), "|")
                    
                    AddFavorit cnt(i), FixPath1(temp(0)) & " - " & FixPath2(temp(1))
                End If
            Next i
        End If
    End With
End Sub
Public Sub CreateSetting()
    Dim TmpSet As String
    Dim Alamat As String
    Dim AlamatNew As String
    Dim pos As Long
    Dim Pot() As String
    
    AlamatNew = GetText(FrmUtama.tAlamat)
    pos = InStr(Lawas, AlamatNew)
    If pos = 0 Then
        Lawas = Lawas & AlamatNew & "|"
        Pot = Split(Lawas, "|")
        '// Hanya menampung maximal 4 alamat
        '// lebih dari itu hapus bagian pertama
        If UBound(Pot) >= 5 Then
            Lawas = Mid$(Lawas, Len(Pot(0)) + 2)
        End If
    End If
    
    With FrmSetting
    
        TmpSet = Head & vbNewLine
        TmpSet = TmpSet & "[CONTROL]" & vbNewLine '// 1
        TmpSet = TmpSet & "Method=1" & vbNewLine
        TmpSet = TmpSet & "Compress=1" & vbNewLine
        TmpSet = TmpSet & "Archive=1" & vbNewLine
        TmpSet = TmpSet & "Log=" & .cLog.Value & vbNewLine
        
        TmpSet = TmpSet & "[LISTVIEW]" & vbNewLine '// 6
        TmpSet = TmpSet & "TipeView=" & GetView & vbNewLine
        TmpSet = TmpSet & "GridLine=" & .cShowGrid.Value & vbNewLine
        TmpSet = TmpSet & "Fullselect=" & .cFullselect.Value & vbNewLine
        
        TmpSet = TmpSet & "[TOOLBAR]" & vbNewLine '// 10
        TmpSet = TmpSet & "ViewMain=" & .cToolbar(0).Value & vbNewLine
        TmpSet = TmpSet & "ButtonUp=" & .cToolbar(1).Value & vbNewLine
        TmpSet = TmpSet & "ViewAdress=" & .cToolbar(2).Value & vbNewLine
        TmpSet = TmpSet & "ShowButton=" & .cToolbar(3).Value & vbNewLine
        TmpSet = TmpSet & "LargeButton=" & .cToolbar(4).Value & vbNewLine
        TmpSet = TmpSet & "LockToolbar=" & .cToolbar(5).Value & vbNewLine
        
        TmpSet = TmpSet & "[BUTTON]" & vbNewLine '// 17
        TmpSet = TmpSet & "Add=" & .cButton(0).Value & vbNewLine
        TmpSet = TmpSet & "ExtractTo=" & .cButton(1).Value & vbNewLine
        TmpSet = TmpSet & "Test=" & .cButton(2).Value & vbNewLine
        TmpSet = TmpSet & "View=" & .cButton(3).Value & vbNewLine
        TmpSet = TmpSet & "Delete=" & .cButton(4).Value & vbNewLine
        TmpSet = TmpSet & "Find=" & .cButton(5).Value & vbNewLine
        TmpSet = TmpSet & "Print=" & .cButton(6).Value & vbNewLine
        TmpSet = TmpSet & "Wizard=" & .cButton(7).Value & vbNewLine
        TmpSet = TmpSet & "Convert=" & .cButton(8).Value & vbNewLine
        TmpSet = TmpSet & "Info=" & .cButton(9).Value & vbNewLine
        TmpSet = TmpSet & "Exit=" & .cButton(10).Value & vbNewLine
        TmpSet = TmpSet & "Repair=" & .cButton(11).Value & vbNewLine
        TmpSet = TmpSet & "Extract=" & .cButton(12).Value & vbNewLine
        TmpSet = TmpSet & "VirusScan=" & .cButton(13).Value & vbNewLine
        TmpSet = TmpSet & "Commant=" & .cButton(14).Value & vbNewLine
        TmpSet = TmpSet & "Protect=" & .cButton(15).Value & vbNewLine
        TmpSet = TmpSet & "Lock=" & .cButton(16).Value & vbNewLine
        TmpSet = TmpSet & "SFX=" & .cButton(17).Value & vbNewLine
        TmpSet = TmpSet & "Report=" & .cButton(18).Value & vbNewLine
        
        TmpSet = TmpSet & "[TEMPORARY]" & vbNewLine
        TmpSet = TmpSet & .tTemp.Text & vbNewLine
        
        TmpSet = TmpSet & "[Favorits1]" & vbNewLine
        TmpSet = TmpSet & Lawas & vbNewLine
        
        TmpSet = TmpSet & "[Favorits2]" & vbNewLine
        TmpSet = TmpSet & TemPFavo & vbNewLine
        
        TmpSet = TmpSet & "[Find1]" & vbNewLine
        
        TmpSet = TmpSet & "[Find2]" & vbNewLine
        
        TmpSet = TmpSet & "[END]" & vbNewLine
    End With
    
    Alamat = App.Path & "\Settings.txt"
    If PathFileExists(StrPtr(Alamat)) = 1 Then Call HapusFile(Alamat)
    
    Open Alamat For Binary As #1
        Put #1, , TmpSet
    Close #1
End Sub
Public Function FixPath1(Alamat As String) As String
    If Len(Alamat) > 13 Then
        FixPath1 = "..." & Right$(Alamat, 10)
    Else
        FixPath1 = Alamat
    End If
End Function
Public Function FixPath2(Alamat As String) As String
    Dim temp As String
    Dim x() As String
    Dim i As Integer
    
    temp = ""
    If Len(Alamat) > 10 Then
        x = Split(Alamat, "\")
        For i = 0 To UBound(x) - 1
            temp = temp & ".." & Right$(x(i), 4) & "\"
        Next i
        FixPath2 = temp
    Else
        FixPath2 = Alamat
    End If
End Function
Public Function GetView() As String
    Dim i As Byte
    
    With FrmSetting
        For i = 0 To 4
            If .OptView(i).Value Then
                GetView = Str$(i)
                Exit For
            End If
        Next i
    End With
End Function
Public Sub SetView(ByVal Value As Byte)
    
    With FrmSetting
        .OptView(Value).Value = 1
    End With
End Sub


