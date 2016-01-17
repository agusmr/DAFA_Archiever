VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form frmProses 
   BackColor       =   &H80000004&
   Caption         =   "Proses"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   3720
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
   ScaleHeight     =   3270
   ScaleWidth      =   3720
   StartUpPosition =   2  'CenterScreen
   Begin DAFA_Component.ucFrame ucFrame1 
      Height          =   1395
      Left            =   30
      Top             =   30
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2461
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin VB.TextBox TAlamat 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   180
         Width           =   3315
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status         : "
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
         TabIndex        =   9
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblFile 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah File       :"
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
         TabIndex        =   8
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label lblFolder 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Jumlah Folder  :"
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
         TabIndex        =   7
         Top             =   1020
         Width           =   1140
      End
   End
   Begin DAFA_Component.ucFrame ucFrame2 
      Height          =   1305
      Left            =   30
      Top             =   1320
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   2302
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin VB.Label lblPros 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proses   :"
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
         Left            =   180
         TabIndex        =   5
         Top             =   900
         Width           =   675
      End
      Begin VB.Label lblUkuran 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ukuran  :"
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
         Left            =   180
         TabIndex        =   4
         Top             =   480
         Width           =   660
      End
      Begin VB.Label lblWaktu 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ": _"
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
         Left            =   810
         TabIndex        =   3
         Top             =   690
         Width           =   195
      End
      Begin VB.Label lblJum 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "File        : "
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
         Left            =   180
         TabIndex        =   2
         Top             =   210
         Width           =   705
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Waktu   "
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
         Left            =   180
         TabIndex        =   1
         Top             =   690
         Width           =   600
      End
   End
   Begin DAFA_Component.ucProgressBar p 
      Height          =   165
      Left            =   30
      Top             =   2640
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   291
   End
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
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
      Left            =   2700
      TabIndex        =   0
      Top             =   2880
      Width           =   1005
   End
End
Attribute VB_Name = "frmProses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Jum As Long
Private Sub cmdKeluar_Click()
    Call VbCloseHandle(hFileDibuat)
    HapusFile xSimpan
    End
End Sub

Private Sub TAlamat_Change()
    On Error Resume Next
    
    Jum = Jum + 1
    lblJum.Caption = "File        : " & CStr(Jum)
    p.Value = Left(CStr(Round(Jum / JumFile * 100, 2)), 5)
    lblPros.Caption = "Proses   : " & p.Value & " %"
    
    If Jum >= JumFile Then
        Jum = 0
    End If
    
    myDoEvents
End Sub

Private Sub ucFrame2_Click()

End Sub
