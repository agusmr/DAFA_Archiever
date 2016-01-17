VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form FrmPilih 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5730
   ClientLeft      =   45
   ClientTop       =   450
   ClientWidth     =   5985
   DrawMode        =   16  'Merge Pen
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   5985
   StartUpPosition =   2  'CenterScreen
   Begin DAFA_Component.ucFrame FUmum 
      Height          =   4455
      Left            =   120
      Top             =   480
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   7858
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin VB.CommandButton cmdSFX 
         Caption         =   "Atur SFX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3000
         TabIndex        =   32
         Top             =   2760
         Width           =   2535
      End
      Begin VB.CommandButton cmdSimpan 
         Caption         =   "..."
         Height          =   315
         Left            =   4740
         TabIndex        =   8
         Top             =   660
         Width           =   675
      End
      Begin VB.CommandButton cmdAlamat 
         Caption         =   "..."
         Height          =   315
         Left            =   4740
         TabIndex        =   7
         Top             =   210
         Width           =   675
      End
      Begin VB.CommandButton CmdAPesan 
         Caption         =   "Atur pesan"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2940
         TabIndex        =   4
         Top             =   3660
         Width           =   2655
      End
      Begin VB.CommandButton CmdAPassword 
         Caption         =   "Atur Password"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2940
         TabIndex        =   3
         Top             =   3210
         Width           =   2655
      End
      Begin DAFA_Component.ucFrame ucFrame2 
         Height          =   1485
         Left            =   180
         Top             =   1050
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2619
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.OptionButton OptZip 
            BackColor       =   &H80000004&
            Caption         =   "Format Zip"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   270
            TabIndex        =   30
            Top             =   450
            Width           =   1275
         End
         Begin VB.OptionButton OpGus 
            BackColor       =   &H80000004&
            Caption         =   "Format Gus"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   270
            MaskColor       =   &H00FFFFFF&
            TabIndex        =   11
            Top             =   120
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton OpRaX 
            BackColor       =   &H80000004&
            Caption         =   "Format RaX"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   270
            TabIndex        =   10
            Top             =   780
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H80000004&
            Caption         =   "Format RaY"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   270
            TabIndex        =   9
            Top             =   1110
            Width           =   1275
         End
      End
      Begin DAFA_Component.ucFrame ucFrame8 
         Height          =   1485
         Left            =   2940
         Top             =   1050
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   2619
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.CheckBox CSolid 
            Caption         =   "Create solid archive"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   31
            Top             =   1020
            Width           =   2025
         End
         Begin VB.CheckBox ckLock 
            Caption         =   "Lock archive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   660
            Width           =   1665
         End
         Begin VB.CheckBox CkSFX 
            Caption         =   "Create SFX archive"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   2145
         End
      End
      Begin DAFA_Component.ucFrame ucFrame4 
         Height          =   1455
         Left            =   180
         Top             =   2610
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2566
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.ComboBox cMethod 
            BeginProperty Font 
               Name            =   "DokChampa"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "FrmPilih.frx":0000
            Left            =   210
            List            =   "FrmPilih.frx":0002
            TabIndex        =   15
            Text            =   "Deflate"
            Top             =   270
            Width           =   1755
         End
         Begin VB.ComboBox ComRatio 
            BeginProperty Font 
               Name            =   "DokChampa"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            ItemData        =   "FrmPilih.frx":0004
            Left            =   210
            List            =   "FrmPilih.frx":0006
            TabIndex        =   14
            Text            =   "Best"
            Top             =   750
            Width           =   1755
         End
      End
      Begin DAFA_Component.ucComboBoxEx tSimpanX 
         Height          =   375
         Left            =   1770
         TabIndex        =   28
         Top             =   660
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   423
         BeginProperty Fnt {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
      End
      Begin DAFA_Component.ucComboBoxEx tAlamat 
         Height          =   375
         Left            =   1770
         TabIndex        =   29
         Top             =   240
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   423
         BeginProperty Fnt {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Simpan Arsip Ke  :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   6
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File yang di Arsip :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   5
         Top             =   300
         Width           =   1320
      End
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   405
      Left            =   3840
      TabIndex        =   1
      Top             =   5040
      Width           =   1935
   End
   Begin VB.CommandButton cmdBuat 
      Caption         =   "Ok"
      Height          =   405
      Left            =   1560
      TabIndex        =   0
      Top             =   5040
      Width           =   1935
   End
   Begin DAFA_Component.ucFrame FPassword 
      Height          =   4425
      Left            =   120
      Top             =   480
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   7805
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin DAFA_Component.ucFrame ucFrame1 
         Height          =   2355
         Left            =   570
         Top             =   150
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   4154
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.TextBox TPastikan 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   390
            PasswordChar    =   "*"
            TabIndex        =   23
            Top             =   1200
            Width           =   3465
         End
         Begin VB.TextBox TPassword 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   222
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            IMEMode         =   3  'DISABLE
            Left            =   390
            PasswordChar    =   "*"
            TabIndex        =   22
            Top             =   510
            Width           =   3465
         End
         Begin VB.CommandButton cmdTes 
            Caption         =   "Tes Password"
            Height          =   315
            Left            =   390
            TabIndex        =   21
            Top             =   1740
            Width           =   1515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masukkan Password"
            Height          =   195
            Left            =   420
            TabIndex        =   25
            Top             =   270
            Width           =   1440
         End
         Begin VB.Label lblPastikan 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Masukkan Password kembali untuk pencocokan"
            Height          =   195
            Left            =   390
            TabIndex        =   24
            Top             =   930
            Width           =   3360
         End
      End
      Begin DAFA_Component.ucFrame ucFrame3 
         Height          =   1245
         Left            =   570
         Top             =   2880
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   2196
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.CheckBox CShow 
            BackColor       =   &H80000004&
            Caption         =   "Tampilkan Password"
            Height          =   195
            Left            =   390
            TabIndex        =   27
            Top             =   360
            Width           =   1785
         End
         Begin VB.CheckBox CEncript 
            BackColor       =   &H80000004&
            Caption         =   "Encrip Nama"
            Height          =   195
            Left            =   390
            TabIndex        =   26
            Top             =   780
            Width           =   1365
         End
      End
   End
   Begin DAFA_Component.ucFrame fCommand 
      Height          =   4425
      Left            =   120
      Top             =   480
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   7805
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin VB.CommandButton cmdBrow 
         Caption         =   "Browse..."
         Height          =   375
         Left            =   4560
         TabIndex        =   18
         Top             =   420
         Width           =   945
      End
      Begin VB.ComboBox cAlamat 
         Height          =   315
         Left            =   150
         TabIndex        =   17
         Top             =   480
         Width           =   4155
      End
      Begin VB.TextBox tComment 
         Height          =   3225
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   16
         Top             =   1080
         Width           =   5325
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Load a comment from the file"
         Height          =   195
         Left            =   180
         TabIndex        =   20
         Top             =   180
         Width           =   2085
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Enter a comment manually"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   810
         Width           =   1890
      End
   End
   Begin DAFA_Component.ucTabStrip Tab1 
      Height          =   5625
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   5925
      _ExtentX        =   10451
      _ExtentY        =   9922
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
End
Attribute VB_Name = "FrmPilih"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CkShow_Click()

End Sub



Private Sub cmdAlamat_Click()
    FrmBrow.Show 1
    tAlamat.Text = TesSlash(tSimpan)
End Sub

Private Sub CmdAPassword_Click()
        FUmum.Visible = False
        FPassword.Visible = True
End Sub

Private Sub cmdBuat_Click()
    Terbuka = True
    
    If GetText(tSimpanX) = "" Then
        cmdSimpan_Click
    End If
    
    If TPassword.Text <> TPastikan.Text And CShow.Value = 0 Then
        MsgBox "Password tidak cocok !!", vbCritical
    Else
        
        pass = TPassword.Text
        DiEncrip = CEncript.Value
        DiLock = ckLock.Value
        DiPassword = Len(pass)
        Pesan = tComment.Text
        DiPesan = Len(Pesan)
        MethodC = cMethod.ListIndex + 1
        RatioC = GetRatio(ComRatio.List(ComRatio.ListIndex))
        
        If InStr(GetText(tAlamat), "?") Then
            BuatArchive TesSlash(Alamat), xSimpan, frmProses.lblStatus
        Else
            BuatArchive TesSlash(GetText(tAlamat)), GetText(tSimpanX), frmProses.lblStatus
        End If
    End If
End Sub
Private Function GetRatio(ByVal Tipe As String) As Long
    Select Case Tipe
        Case "No Compression"
            GetRatio = 0
        Case "Low"
            GetRatio = 4
        Case "Standard"
            GetRatio = 6
        Case "High"
            GetRatio = 7
        Case "Best"
            GetRatio = 9
    End Select
End Function

Private Sub cmdKeluar_Click()
    If Terbuka Then
        Unload Me
    Else
        End
    End If
End Sub
Private Sub cmdSimpan_Click()
    With FrmUtama.c
        '.FileNames = ""
        .FileFilter = "Darma Archive|*.gus"
        .ShowSave
        If .FileName <> "" Then
            tSimpanX.Text = StripNulls(CStr(.FileName))
        End If
    End With
End Sub
Private Sub cmdTes_Click()
    If TPassword.Text <> TPastikan.Text And CShow.Value = 0 Then
        MsgBox "Password tidak cocok !!", vbCritical
    Else
        MsgBox "Password cocok !!", vbInformation
    End If
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub cMethod_Click()
    Select Case cMethod.ListIndex
        Case 0
            ComRatio.Enabled = True
            '    Z_NO_COMPRESSION = 0
            '    Z_BEST_SPEED = 1
            '    Z_BEST_COMPRESSION = 9
            '    Z_DEFAULT_COMPRESSION = -1
            ComRatio.List(0) = "No Compression"
            ComRatio.List(1) = "Low"
            ComRatio.List(2) = "Standard"
            ComRatio.List(3) = "High"
            ComRatio.List(4) = "Best"
        Case 1
            ComRatio.Enabled = True
        Case 2
            ComRatio.Enabled = False
        Case 3
            ComRatio.Enabled = False
        Case 4
            ComRatio.Enabled = False
    End Select
End Sub


Private Sub CShow_Click()
    If CShow.Value = 1 Then
        lblPastikan.Visible = False
        TPassword.PasswordChar = ""
        TPastikan.Visible = False
        cmdTes.Visible = False
    Else
        lblPastikan.Visible = True
        TPassword.PasswordChar = "*"
        TPastikan.Visible = True
        cmdTes.Visible = True
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    KeepOnTop Me, True
    
    ComRatio.AddItem "No Compression"
    ComRatio.AddItem "Low"
    ComRatio.AddItem "High"
    ComRatio.AddItem "Best"
    
    cMethod.AddItem "Deflate"
    cMethod.AddItem "bzip2"
    cMethod.AddItem "LZMA"
    cMethod.AddItem "APlib"
    cMethod.AddItem "fLz"
    cMethod.ListIndex = 0
    ComRatio.ListIndex = 2
    
    With Tab1
        .Tabs.Add "Umum"
        .Tabs.Add "SFX"
        .Tabs.Add "Password"
        .Tabs.Add "Pesan"
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Terbuka Then
        Unload Me
    Else
        End
    End If
End Sub

Private Sub FUmum_Click()

End Sub

Private Sub FPassword_Click()

End Sub

Private Sub Tab1_Click(ByVal oTab As DAFA_Component.cTab)
    Select Case oTab.Index
        Case 1
            FUmum.Visible = True
        Case 3
            FUmum.Visible = False
            FPassword.Visible = True
            TPassword.SetFocus
        Case 4
            FUmum.Visible = False
            FPassword.Visible = False
            fCommand.Visible = True
    End Select
End Sub

Private Sub ucFrame10_Resize()

End Sub

Private Sub ucFrame9_Resize()

End Sub

Private Sub TAlamat_Change()

End Sub

Private Sub TPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBuat_Click

End Sub

Private Sub TPastikan_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdBuat_Click
End Sub
