VERSION 5.00
Begin VB.Form FrmBrow 
   Caption         =   "Extract path"
   ClientHeight    =   6210
   ClientLeft      =   1320
   ClientTop       =   1710
   ClientWidth     =   7545
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
   ScaleHeight     =   6210
   ScaleWidth      =   7545
   Begin DaChiVleR.DirCtl DirCtl1 
      Height          =   4905
      Left            =   3330
      TabIndex        =   4
      Top             =   540
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   8652
   End
   Begin VB.CommandButton cmdNewF 
      Caption         =   "New folder"
      Height          =   345
      Left            =   5970
      TabIndex        =   3
      Top             =   90
      Width           =   1395
   End
   Begin VB.ComboBox cAlamat 
      Height          =   315
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   5625
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5910
      TabIndex        =   1
      Top             =   5610
      Width           =   1485
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   4200
      TabIndex        =   0
      Top             =   5610
      Width           =   1485
   End
End
Attribute VB_Name = "FrmBrow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Yes As Boolean
Private Sub cmdCancel_Click()
    tSimpan = ""
    Unload Me
End Sub

Private Sub cmdNewF_Click()
    Dim NewF As String
    Dim FullPath As String
    
    NewF = InputBox("Masukkan nama folder", "Darma File Archive", "New Folder")
    FullPath = tSimpan & TesSlash(NewF)
    
    If NewF <> "" Then
        BuatFolder FullPath
        DirCtl1.LoadTreeView
        DirCtl1.Tampilkan FullPath
    End If
End Sub

Private Sub cmdOk_Click()
    Yes = True
    Unload Me
End Sub

Private Sub DirCtl1_DirPath(ByVal spath As String)
    tSimpan = spath
    cAlamat.text = tSimpan
End Sub

Private Sub Form_Load()
    Yes = False
    KeepOnTop Me, True
    DirCtl1.LoadTreeView
    DirCtl1.Tampilkan tSimpan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Yes = False Then tSimpan = ""
    Unload Me
End Sub
