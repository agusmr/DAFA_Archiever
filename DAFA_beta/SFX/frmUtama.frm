VERSION 5.00
Begin VB.Form frmUtama 
   Caption         =   "DAFA Installer"
   ClientHeight    =   5130
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6720
   Icon            =   "frmUtama.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6720
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   0
      ScaleHeight     =   4185
      ScaleWidth      =   2025
      TabIndex        =   5
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   2160
      TabIndex        =   4
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox tInfo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   2160
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   0
      Width           =   4455
   End
   Begin VB.CommandButton cmdBrow 
      Caption         =   "..."
      Height          =   375
      Left            =   6120
      TabIndex        =   2
      Top             =   3840
      Width           =   495
   End
   Begin VB.TextBox tPath 
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   3840
      Width           =   3855
   End
   Begin VB.CommandButton cmdInstall 
      Caption         =   "Install"
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   4440
      Width           =   1935
   End
End
Attribute VB_Name = "frmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Path As String
Dim loc As String
Private Sub cmdBrow_Click()

    Path = BrowseForFolder(Me, "", "Select Path to Install DAFA !")
    tPath.Text = Path
End Sub

Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdInstall_Click()
    If Path <> "" Then
        loc = TesSlash(Path) & "DAFA\"
        If (ExtractArchive(App.Path & "\" & App.EXEName & ".exe", Path)) Then
            Panggil "Regsvr32", "/s /c " & loc & "Componen.ocx", 5
            
            Panggil loc & "DAFA.exe", "", 5
            
        Else
           MsgBox "Proses Ekstraksi File Gagal !", vbCritical
        End If
    Else
        MsgBox "Pilih dulu Lokasi untuk install program !", vbCritical
    End If
End Sub

Private Sub Form_Load()
Panggil "Regsvr32", "/c " & "\'D:\DAFA Next Generation\Edit - Copy\Componen.ocx\'", 5
End Sub
