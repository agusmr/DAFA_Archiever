VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form FrmAbout 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About DaChiVleR"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6090
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin DAFA_Component.ucFrame ucFrame1 
      Height          =   1065
      Left            =   60
      Top             =   1020
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   1879
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   90
         X2              =   2790
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DaChiVleR Beta 0.1"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   150
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Copyright @ 2010 - 2011 DarmaSoft"
         Height          =   195
         Left            =   90
         TabIndex        =   5
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "For 32-bit Windows Development"
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   390
         Width           =   2385
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   4890
      TabIndex        =   3
      Top             =   6660
      Width           =   1125
   End
   Begin VB.Timer t1 
      Interval        =   100
      Left            =   4620
      Top             =   3000
   End
   Begin VB.PictureBox PicSpon 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Height          =   4485
      Left            =   60
      ScaleHeight     =   4425
      ScaleWidth      =   5895
      TabIndex        =   0
      Top             =   2130
      Width           =   5955
      Begin VB.PictureBox PicLogo2 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   705
         Left            =   60
         ScaleHeight     =   705
         ScaleWidth      =   975
         TabIndex        =   2
         Top             =   60
         Width           =   975
      End
      Begin VB.TextBox tSpon 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8715
         Left            =   0
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3240
         Width           =   5835
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H8000000C&
      X1              =   90
      X2              =   6000
      Y1              =   2040
      Y2              =   2040
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   945
      Left            =   60
      Picture         =   "FrmAbout.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   5985
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = FrmUtama.Icon
    PicLogo2.Top = tSpon.Top - 50
    PicLogo2.Picture = FrmUtama.Icon
    tSpon.text = "DARMA FILE ARCHIVER" & vbCrLf & _
                "DaChiVleR" & vbCrLf & vbCrLf & _
                "===========================================" & vbCrLf & vbCrLf & _
                "Code Developer" & vbCrLf & _
                "Agus Minanur Rohman" & vbCrLf & vbCrLf & _
                "Support" & vbCrLf & vbCrLf & _
                "AgrotiX Packing Algorithm" & vbCrLf & _
                "Agus M.R." & vbCrLf & vbCrLf & _
                "Algorithm RC4" & vbCrLf & _
                "Ron Rivest" & vbCrLf & vbCrLf & _
                "zlib data compression library" & vbCrLf & _
                "Jean-loup Gailly & Mark Adler" & vbCrLf & vbCrLf & _
                "RAR decompression library" & vbCrLf & _
                "Alexander Roshal" & vbCrLf & vbCrLf & _
                "LZMA library" & vbCrLf & _
                "Igor Pavlov" & vbCrLf & vbCrLf & _
                "Bzib2 library" & vbCrLf & _
                "Arnout de Vries, Relevant Soft- & Mindware" & vbCrLf & vbCrLf & _
                "Hashing Crc32" & vbCrLf & _
                "Noel A. Dacara" & vbCrLf & _
                "Fredrik Qvarfort" & vbCrLf & vbCrLf & _
                "Hashing Crc16" & vbCrLf & _
                "Noel A. Dacara" & vbCrLf & _
                "Fredrik Qvarfort"
End Sub

Private Sub t1_Timer()

    If tSpon.Height + tSpon.Top >= 0 Then
        tSpon.Top = tSpon.Top - 50
        PicLogo2.Top = tSpon.Top - 50
    Else
        tSpon.Top = PicSpon.Height
    End If
End Sub

