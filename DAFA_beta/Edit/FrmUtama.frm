VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form FrmUtama 
   Caption         =   "ArchiGus"
   ClientHeight    =   6315
   ClientLeft      =   165
   ClientTop       =   480
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmUtama.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmUtama.frx":1CFA
   ScaleHeight     =   6315
   ScaleWidth      =   10365
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic4 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9600
      Picture         =   "FrmUtama.frx":39F4
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   33
      Top             =   6840
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.ListBox l 
      Height          =   255
      Left            =   6480
      TabIndex        =   32
      Top             =   4890
      Visible         =   0   'False
      Width           =   435
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   18
      Left            =   4290
      Picture         =   "FrmUtama.frx":3F7E
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   31
      Top             =   8580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   16
      Left            =   2670
      Picture         =   "FrmUtama.frx":4C18
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   30
      Top             =   8580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   17
      Left            =   3480
      Picture         =   "FrmUtama.frx":58B2
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   29
      Top             =   8580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   11
      Left            =   9120
      Picture         =   "FrmUtama.frx":654C
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   28
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   10
      Left            =   8310
      Picture         =   "FrmUtama.frx":71E6
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   27
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   8
      Left            =   6750
      Picture         =   "FrmUtama.frx":7E80
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   26
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   6
      Left            =   5070
      Picture         =   "FrmUtama.frx":8B1A
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   25
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin DAFA_Component.ucComboBoxEx tAlamat 
      Height          =   330
      Left            =   570
      TabIndex        =   24
      Top             =   1020
      Width           =   9705
      _ExtentX        =   17119
      _ExtentY        =   423
      BeginProperty Fnt {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      ExtUI           =   -1  'True
   End
   Begin DAFA_Component.ucRebar r 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   1720
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin DAFA_Component.ucStatusBar status 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      Top             =   6015
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   529
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin DAFA_Component.ucListView LVSelect 
      Height          =   615
      Left            =   3960
      TabIndex        =   21
      Top             =   3240
      Visible         =   0   'False
      Width           =   675
      _ExtentX        =   1191
      _ExtentY        =   1085
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin VB.PictureBox PicSplit 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   9960
      ScaleHeight     =   2116.253
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   20
      Top             =   1380
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox PicRe 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8490
      Picture         =   "FrmUtama.frx":97B4
      ScaleHeight     =   270
      ScaleWidth      =   270
      TabIndex        =   18
      Top             =   4620
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.PictureBox Picf 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   7320
      Picture         =   "FrmUtama.frx":9DAE
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   17
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   4080
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   1
      Left            =   990
      Picture         =   "FrmUtama.frx":A678
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   15
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   3
      Left            =   2670
      Picture         =   "FrmUtama.frx":B312
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   14
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   4
      Left            =   3480
      Picture         =   "FrmUtama.frx":BFAC
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   13
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   2
      Left            =   1830
      Picture         =   "FrmUtama.frx":CC46
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   12
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   0
      Left            =   90
      Picture         =   "FrmUtama.frx":D8E0
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   11
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   5
      Left            =   4260
      Picture         =   "FrmUtama.frx":E57A
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   10
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   7
      Left            =   5940
      Picture         =   "FrmUtama.frx":F214
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   9
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   9
      Left            =   7530
      Picture         =   "FrmUtama.frx":FEAE
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   8
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   13
      Left            =   120
      Picture         =   "FrmUtama.frx":10B48
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   7
      Top             =   8580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   14
      Left            =   990
      Picture         =   "FrmUtama.frx":117E2
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   6
      Top             =   8580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   15
      Left            =   1830
      Picture         =   "FrmUtama.frx":1247C
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   5
      Top             =   8580
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   570
      Index           =   12
      Left            =   9990
      Picture         =   "FrmUtama.frx":13116
      ScaleHeight     =   570
      ScaleWidth      =   720
      TabIndex        =   4
      Top             =   7830
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8700
      Picture         =   "FrmUtama.frx":13DB0
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   3
      Top             =   6870
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7380
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   2
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   6870
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox pic3 
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8190
      Picture         =   "FrmUtama.frx":1433A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   6960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Timer Tim1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5670
      Top             =   5010
   End
   Begin VB.PictureBox picSplitter 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      FillColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   3390
      ScaleHeight     =   2116.253
      ScaleMode       =   0  'User
      ScaleWidth      =   780
      TabIndex        =   0
      Top             =   1380
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.PictureBox PicL 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "DokChampa"
         Size            =   8.25
         Charset         =   222
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   8400
      ScaleHeight     =   32
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   32
      TabIndex        =   16
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   5490
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.TextBox tPesan 
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   10080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   19
      Top             =   1380
      Visible         =   0   'False
      Width           =   5205
   End
   Begin DAFA_Component.ucToolbar t 
      Height          =   315
      Index           =   1
      Left            =   7830
      Top             =   5070
      Width           =   435
      _ExtentX        =   767
      _ExtentY        =   556
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin DAFA_Component.ucToolbar t 
      Height          =   285
      Index           =   0
      Left            =   7830
      Top             =   5100
      Visible         =   0   'False
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   503
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin DAFA_Component.ucTreeView TV 
      Height          =   4845
      Left            =   30
      TabIndex        =   23
      Top             =   1530
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   8546
   End
   Begin DAFA_Component.ucListView LVRead 
      Height          =   1455
      Left            =   3480
      TabIndex        =   22
      Top             =   1440
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   2566
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      Style           =   16
      StyleEx         =   32
      OleDrop         =   -1  'True
   End
   Begin DAFA_Component.c c 
      Left            =   4200
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      FileFlags       =   2621444
      FolderFlags     =   323
      FileCustomFilter=   "FrmUtama.frx":148C4
      FileDefaultExtension=   "FrmUtama.frx":148E4
      FileFilter      =   "FrmUtama.frx":14904
      FileOpenTitle   =   "FrmUtama.frx":1494C
      FileSaveTitle   =   "FrmUtama.frx":14984
      FolderMessage   =   "FrmUtama.frx":149BC
   End
   Begin VB.Image ImgSplit 
      Height          =   4905
      Left            =   9960
      MousePointer    =   9  'Size W E
      Top             =   1380
      Width           =   60
   End
   Begin VB.Image imgSplitter 
      Height          =   4905
      Left            =   3390
      MousePointer    =   9  'Size W E
      Top             =   1380
      Width           =   60
   End
   Begin VB.Menu File 
      Caption         =   "File"
      Begin VB.Menu BukaFileArchive 
         Caption         =   "Open archive"
         Shortcut        =   ^O
      End
      Begin VB.Menu SaveCopy 
         Caption         =   "Save archive copy as .."
      End
      Begin VB.Menu SetPasdd 
         Caption         =   "Set default password"
      End
      Begin VB.Menu fsafd 
         Caption         =   "-"
      End
      Begin VB.Menu Copyfils 
         Caption         =   "Copy files to clipboard"
      End
      Begin VB.Menu Past 
         Caption         =   "Paste files from clipboard"
      End
      Begin VB.Menu xggc 
         Caption         =   "-"
      End
      Begin VB.Menu SelAll 
         Caption         =   "Select all"
      End
      Begin VB.Menu SelGrop 
         Caption         =   "Select group"
      End
      Begin VB.Menu DesGrop 
         Caption         =   "Deselect group"
      End
      Begin VB.Menu InvSel 
         Caption         =   "Invert selection"
      End
      Begin VB.Menu xghh 
         Caption         =   "-"
      End
      Begin VB.Menu BuatArchiveBaru 
         Caption         =   "New archive"
         Begin VB.Menu DalamFolder 
            Caption         =   "By Folder"
         End
         Begin VB.Menu Pisah1 
            Caption         =   "-"
         End
         Begin VB.Menu SatuFile 
            Caption         =   "Single File"
         End
      End
      Begin VB.Menu zz 
         Caption         =   "-"
      End
      Begin VB.Menu Keluar1 
         Caption         =   "Close current archive"
      End
      Begin VB.Menu Keluar 
         Caption         =   "Exit"
      End
      Begin VB.Menu MnFavo 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu MnFavo 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Tool 
      Caption         =   "Tools"
      Begin VB.Menu Wiz 
         Caption         =   "Wizard"
      End
      Begin VB.Menu zxdg 
         Caption         =   "-"
      End
      Begin VB.Menu ScanArc 
         Caption         =   "Scan archive for viruses"
      End
      Begin VB.Menu ConvArc 
         Caption         =   "Convert archives"
      End
      Begin VB.Menu RepArc 
         Caption         =   "Repair archive"
      End
      Begin VB.Menu ConvToSFX 
         Caption         =   "Convert archive to SFX"
      End
      Begin VB.Menu dxgd 
         Caption         =   "-"
      End
      Begin VB.Menu FindFile 
         Caption         =   "Find files"
      End
      Begin VB.Menu ShowInfor 
         Caption         =   "Show information"
      End
      Begin VB.Menu GeneratRep 
         Caption         =   "Generate report"
      End
   End
   Begin VB.Menu Favor 
      Caption         =   "Favorites"
      Begin VB.Menu AddFavor 
         Caption         =   "Add to favorites"
      End
      Begin VB.Menu OrgFavor 
         Caption         =   "Organize favorites"
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   "-"
         Index           =   0
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFavorites 
         Caption         =   ""
         Index           =   1
         Visible         =   0   'False
      End
   End
   Begin VB.Menu Commans 
      Caption         =   "Commands"
      Begin VB.Menu SelectAll 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu sd 
         Caption         =   "-"
      End
      Begin VB.Menu TambahFile 
         Caption         =   "Add Files to Archive"
         Shortcut        =   ^T
      End
      Begin VB.Menu AddFol 
         Caption         =   "Add Folder to Archive"
         Shortcut        =   ^D
      End
      Begin VB.Menu ExtractKe 
         Caption         =   "Extact to the specified Folder"
         Shortcut        =   ^F
      End
      Begin VB.Menu testArchF 
         Caption         =   "Test archived file"
      End
      Begin VB.Menu BukaFilenya 
         Caption         =   "View files"
         Shortcut        =   ^B
      End
      Begin VB.Menu BukaNotepad 
         Caption         =   "View files with notepad"
      End
      Begin VB.Menu HapusFile 
         Caption         =   "Delete Files"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu UbahNama 
         Caption         =   "Rename Files"
         Shortcut        =   ^S
      End
      Begin VB.Menu PrintFF 
         Caption         =   "Print file"
      End
      Begin VB.Menu RepArca 
         Caption         =   "Repair archive"
      End
      Begin VB.Menu ExWInfo 
         Caption         =   "Extract without confirmation"
      End
      Begin VB.Menu AddArcCom 
         Caption         =   "Add archive command"
      End
      Begin VB.Menu ProtArcFromD 
         Caption         =   "Protect archive from demage"
      End
      Begin VB.Menu LockArc 
         Caption         =   "Lock archive"
      End
      Begin VB.Menu InformasiFile 
         Caption         =   "Show Information"
      End
      Begin VB.Menu sds 
         Caption         =   "-"
      End
      Begin VB.Menu SetPassw 
         Caption         =   "Set default password"
         Shortcut        =   ^P
      End
      Begin VB.Menu AdFavo 
         Caption         =   "Add favorites"
      End
      Begin VB.Menu xcs 
         Caption         =   "-"
      End
      Begin VB.Menu CreteFol 
         Caption         =   "Create new folder"
      End
      Begin VB.Menu ViewAs 
         Caption         =   "View as"
         Begin VB.Menu ModIco 
            Caption         =   "icons"
         End
         Begin VB.Menu ModDetail 
            Caption         =   "Detail"
         End
         Begin VB.Menu SmlIcon 
            Caption         =   "Small Icon"
         End
         Begin VB.Menu ModList 
            Caption         =   "List"
         End
         Begin VB.Menu ModTit 
            Caption         =   "Title"
         End
      End
   End
   Begin VB.Menu OptionX 
      Caption         =   "Option"
      Begin VB.Menu SettingX 
         Caption         =   "Settings"
      End
      Begin VB.Menu ImEx 
         Caption         =   "Import/Export"
         Begin VB.Menu ImpFF 
            Caption         =   "Import settings from file"
         End
         Begin VB.Menu ExpFF 
            Caption         =   "Export settings to file"
         End
      End
      Begin VB.Menu dfsg 
         Caption         =   "-"
      End
      Begin VB.Menu FileLisX 
         Caption         =   "File list"
         Begin VB.Menu FltFolderV 
            Caption         =   "Flat folders view"
         End
         Begin VB.Menu fddrga 
            Caption         =   "-"
         End
         Begin VB.Menu LvX 
            Caption         =   "List view"
         End
         Begin VB.Menu DetailX 
            Caption         =   "Detail"
         End
      End
      Begin VB.Menu dFT 
         Caption         =   "-"
      End
      Begin VB.Menu ViewLg 
         Caption         =   "View log"
      End
      Begin VB.Menu ClearL 
         Caption         =   "Clear log"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Help"
      Begin VB.Menu Petunjuk 
         Caption         =   "Petunjuk"
      End
      Begin VB.Menu qq 
         Caption         =   "-"
      End
      Begin VB.Menu Home 
         Caption         =   "Home Page"
      End
      Begin VB.Menu ww 
         Caption         =   "-"
      End
      Begin VB.Menu About 
         Caption         =   "Mengenai DaChiVleR"
      End
   End
End
Attribute VB_Name = "FrmUtama"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mbMoving As Boolean
Private Type tagInitCommonControlsEx
   lngSize  As Long
   lngICC   As Long
End Type
Private Declare Sub InitCommonControls Lib "Comctl32" ()
Private Declare Function InitCommonControlsEx Lib "Comctl32" (iCCex As tagInitCommonControlsEx) As Boolean
Private Declare Function MessageBoxW Lib "user32.dll" (ByVal hwnd As Long, ByVal lpText As Long, ByVal lpCaption As Long, ByVal wType As Long) As Long
Private Const OLEDRAG_Listview As Long = &H1234 'DataObject format value
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

'Const sglSplitLimit = 500
Private dragComplate As Boolean
Dim dataku As DataObject
Private Sub AddF_Click()
End Sub

Private Sub About_Click()
    FrmAbout.Show 1, FrmUtama
End Sub

Private Sub AddFavor_Click()
    If Alamat <> "" Then
        If TemPFavo = "" Then
            TemPFavo = Alamat & "|" & FokusNode & "*"
        Else
            TemPFavo = TemPFavo & Alamat & "|" & FokusNode & "*"
        End If
        
        Call CreateSetting
    
        Call LoadSetting
    End If
End Sub

Private Sub AddFol_Click()
    Call TambahkanFolder
End Sub

Private Sub BukaFileArchive_Click()
    Call BukaArsip
End Sub

Private Sub Command3_Click()

End Sub

Private Sub BukaFilenya_Click()
    Call LoadFolder
End Sub


Private Sub c1_OpenFile(ByVal FileName As String)

End Sub

Private Sub c_Click()

End Sub

Private Sub BukaNotepad_Click()
    NamaKey = TesSlash(CariSelect)
    If ApakahFile = True Then
        Dim nama As String
        
        FrmUtama.LVSelect.ListItems.Clear
        nama = CariSelect
        ViewArc nama, 0
    End If
End Sub

Private Sub CreteFol_Click()
    MakeNewFolder Alamat
End Sub

Private Sub DalamFolder_Click()
    FrmBrow.Show 1
    If tSimpan <> "" Then
        FrmPilih.tAlamat.Text = TesSlash(tSimpan)
        Alamat = Left$(tSimpan, Len(tSimpan) - 1)
        xSimpan = Alamat & ".gus"
        FrmPilih.tSimpanX.Text = xSimpan
        FrmPilih.Show
    End If
End Sub

Private Sub ExtractArsip_Click()

End Sub

Private Sub Delsf_Click()
End Sub

Private Sub DetailX_Click()
    LVRead.View = lvwDetails
End Sub

Private Sub ExtractKe_Click()
    Call ExtractAllSpesial
End Sub

Private Sub Exttothespec_Click()
End Sub

Private Sub ExtWW_Click()

End Sub

Private Sub ExWInfo_Click()
    Call ExtractAllArsip
End Sub

Private Sub FindFile_Click()
    Call CariFile
End Sub


Private Sub Form_Initialize()
    On Error Resume Next
        Call InitCommonControls
        Dim iCCex As tagInitCommonControlsEx
            With iCCex
                .lngSize = LenB(iCCex)
                .lngICC = &H1 Or &H2 Or &H20 Or &H4 Or &H10 Or &H200 Or &H400 Or &H8
            End With    'listview;treeview;progressbar;tool-track-tooltip-status;updown;combo;coolbar;tab;
        Call InitCommonControlsEx(iCCex)    'Common control.
End Sub

Private Sub cmd_Click()
    c.ShowFolder
    tAlamat.Text = c.FolderPath
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call LoadPesan
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Bersihkan TempDir
    End
End Sub

Private Sub Pic6_Click()

End Sub

Private Sub GantiSFX_Click()

End Sub

Private Sub HapusFile_Click()
    Call HapusArsip
End Sub


Private Sub Informasi_Click()

End Sub


Private Sub InformasiFile_Click()
    Call ShowInfo
End Sub

Private Sub Keluar_Click()
    End
End Sub

Private Sub Keluar1_Click()
    Call KeluarArsip
End Sub

Private Sub LockArc_Click()
    Call ShowLock
End Sub

Private Sub LVRead_ContextMenu(ByVal x As Single, ByVal y As Single)
    If GetText(tAlamat) <> "" Then
        PopupMenu Commans, 0 Or 2, , , SelectAll
    End If
End Sub

Private Sub LVRead_ItemActivate(ByVal oItem As DAFA_Component.cListItem)
    Call LoadFolder
End Sub


Private Sub LVRead_ItemDrag(ByVal oItem As DAFA_Component.cListItem, ByVal iButton As DAFA_Component.evbComCtlMouseButton)
    LVRead.OLEDrag
    
End Sub


Private Sub LVRead_OLECompleteDrag(Effect As DAFA_Component.evbComCtlOleDropEffect)
    dragComplate = True
    MsgBox "complate"
End Sub

Private Sub LVRead_OLEDragDrop(data As DataObject, Effect As DAFA_Component.evbComCtlOleDropEffect, Button As DAFA_Component.evbComCtlMouseButton, Shift As DAFA_Component.evbComCtlKeyboardState, x As Single, y As Single)
    Dim Fokus As String
    Dim Jum As Long
    Dim nama As String
    Dim G As cListItem
    Dim potong As String
    
    Jum = 0
    If Effect <> vbDropEffectCopy Then
        If data.GetFormat(OLEDRAG_Listview) Then
            
        ElseIf data.GetFormat(vbCFFiles) Then
            With FrmUtama
                .LVSelect.ListItems.Clear
                If Alamat <> "" Then
                    For i = 1 To data.Files.Count
                        If (GetFileAttributesW(StrPtr(data.Files(i))) And vbDirectory) = 0 Then
                            nama = StripNulls(data.Files(i))
                            Set G = .LVSelect.ListItems.Add(, nama)
                            G.SubItem(2).Text = 1
                            G.SubItem(3).Text = AmbilNama(nama)
                        Else
                            nama = data.Files(i)
                            potong = AmbilAlamat(nama)
                            Kumpulkan TesSlash(nama), -1, potong
                        End If
                    Next
                    Jum = .LVSelect.ListItems.Count
                    TambahArchive Alamat, Jum
            
                End If
            End With
        End If
    End If
    
x:

End Sub



Private Sub Status_Click()

End Sub

Private Sub r_Click()

End Sub

Private Sub LVRead_OLESetData(data As DataObject, DataFormat As Integer)
Dim Tempat As String
    Dim nama As String
    Dim JumSelect As Long
    Dim Fokus As String
    Dim KeyState As Integer
    
    JumSelect = MasukSelect
    dragComplate = False
    KeyState = 1
    Do While KeyState
        
        KeyState = GetAsyncKeyState(vbKeyLButton)
        DoEvents
    Loop
            data.SetData , vbCFFiles
            Dim y As Long

            If ApakahFile = True Then
                If JumSelect > 0 Then
                    ExtractFile Alamat, JumSelect, TempDir
                    For i = 1 To JumSelect
                        Tempat = TempDir & AmbilNama(LVSelect.ListItems(i).Text)
                        data.Files.Add Tempat
                    Next i
                End If
            Else
                If JumSelect > 0 Then
                    ExtractSingleFol Alamat, CariSelect, TempDir
                    Tempat = TempDir & CariSelect
                    data.Files.Add Tempat
                End If
            End If
            Me.WindowState = 1
End Sub

Private Sub LVRead_OLEStartDrag(data As DataObject, AllowedEffects As DAFA_Component.evbComCtlOleDropEffect)

    data.SetData , vbCFFiles
    AllowedEffects = vbccOleDropCopy


End Sub

Private Sub MnFavo_Click(Index As Integer)
    Dim Tmp As String
    
    Tmp = MnFavo(Index).Tag
    
    If Tmp <> Alamat Then
        BacaArchive Tmp
    End If
End Sub

Private Sub mnuFavorites_Click(Index As Integer)
    Dim Tmp As String
    Dim Path() As String
    
    Tmp = mnuFavorites(Index).Tag
    Path = Split(Tmp, "|")
    
    If Path(0) <> Alamat Then
        If BacaArchive(Path(0)) Then
            Tampilkan Path(1)
        End If
    End If
    
End Sub

Private Sub ModDetail_Click()
    Call GantiView(lvwDetails)

End Sub

Private Sub ModIco_Click()
    Call GantiView(lvwIcon)
End Sub

Private Sub ModList_Click()
    Call GantiView(lvwList)
End Sub

Private Sub ModTit_Click()
    Call GantiView(lvwTile)

End Sub

Private Sub Pengaturan_Click()

End Sub

Private Sub Picture1_Click()

End Sub

Private Sub RenFF_Click()
End Sub



Private Sub r_ChevronPushed(ByVal oBand As DAFA_Component.cBand, ByVal fLeft As Single, ByVal fTop As Single, ByVal fWidth As Single, ByVal fHeight As Single)
    Dim loToolbar As ucToolbar
    Set loToolbar = oBand.Child
    
    loToolbar.ShowPopup fLeft + fWidth, fTop, fLeft, fTop, fWidth, fHeight, r.Align

End Sub

Private Sub SelAll_Click()
    Call PilihSemua
End Sub

Private Sub SelectAll_Click()
    Call PilihSemua
    LVRead.SetFocus
End Sub

Private Sub SetPass_Click()
End Sub

Private Sub SetPasdd_Click()
    SetPassw_Click
End Sub

Private Sub SetPassw_Click()
    Call SetPassword
End Sub


Private Sub SettingX_Click()
    FrmSetting.Show
End Sub

Private Sub ShowInfor_Click()
    Call ShowInfo
End Sub

Private Sub SmlIcon_Click()
    Call GantiView(lvwSmallIcon)

End Sub

Private Sub status_PanelDblClick(ByVal oPanel As DAFA_Component.cPanel, ByVal iButton As DAFA_Component.evbComCtlMouseButton)
    If oPanel.Index = 1 Then
        Call SetPassword
    End If
End Sub

Private Sub t_ButtonClick(Index As Integer, ByVal Tombol As DAFA_Component.cButton)
    Select Case Tombol.KEY
        Case "Add"
            ' Add
            Call TambahkanFile
            'BukaFileArchive_Click
        Case "Extract to"
            ' Extract to
            Call ExtractAllSpesial
        Case "Test"
            ' Test
            Call TestArchives
        Case "View"
            ' View
            Call LoadFolder
        Case "Delete"
            ' Delete
            Call HapusArsip
        Case "Find"
            ' Find
            Call CariFile
        Case "Wizard"
            ' Wizard
        Case "Info"
            ' Info
            Call ShowInfo
        Case "Exit"
            End
        Case "Extract"
            Call ExtractAllArsip
        Case "VirusScan"
            ' VirusScan
        Case "Comment"
            Call ShowComment
            ' Comment
        Case "Protect"
            ' Protect
            Call ShowProtect
        Case "SFX"
            ' SFX
        Case "Lock"
            Call ShowLock
    End Select
    
    If Index = 1 Then
        Kembali
    End If
End Sub

Private Sub tAlamat_Click()
    
End Sub

Private Sub tAlamat_ListIndexChange()
    Alamat = GetText(tAlamat)
    BacaArchive Alamat
    Form_Resize
End Sub

Private Sub TambahFile_Click()
    Call TambahkanFile
End Sub
Private Sub ImgSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = x + ImgSplit.Left
        If sglPos < sglSplitLimit Then
            PicSplit.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            PicSplit.Left = Me.Width - sglSplitLimit
        Else
            PicSplit.Left = sglPos
        End If
    End If
    ImgSplit.Top = TV.Top
    ImgSplit.Height = TV.Height
End Sub
Private Sub ImgSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With ImgSplit
        PicSplit.Move .Left, .Top, .Width \ 2, .Height + 20
    End With
    PicSplit.Visible = True
    mbMoving = True
End Sub

Private Sub ImgSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls1 PicSplit.Left
    PicSplit.Visible = False
    mbMoving = False
End Sub

Private Sub imgSplitter_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With imgSplitter
        picSplitter.Move .Left, .Top, .Width \ 2, .Height - 20
    End With
    picSplitter.Visible = True
    mbMoving = True
End Sub
Private Sub imgSplitter_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sglPos As Single
    If mbMoving Then
        sglPos = x + imgSplitter.Left
        If sglPos < sglSplitLimit Then
            picSplitter.Left = sglSplitLimit
        ElseIf sglPos > Me.Width - sglSplitLimit Then
            picSplitter.Left = Me.Width - sglSplitLimit
        Else
            picSplitter.Left = sglPos
        End If
    End If
    imgSplitter.Top = TV.Top
    imgSplitter.Height = TV.Height
End Sub
Private Sub imgSplitter_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SizeControls picSplitter.Left
    picSplitter.Visible = False
    mbMoving = False
End Sub
Sub SizeControls(ByVal x As Single)
    On Error Resume Next
    If x < 2700 Then x = 2700
    If x > (Me.Width - 2000) Then x = Me.Width - 2000
    TV.Width = x
    imgSplitter.Left = x + 50
    LVRead.Left = x + 70
    
    If tPesan.Visible = True Then
        LVRead.Width = Me.Width - (tPesan.Width + TV.Width + 300)
    Else
        LVRead.Width = Me.Width - (TV.Width + 380)
    End If
End Sub
Sub SizeControls1(ByVal x As Single)
    On Error Resume Next
    If x < 2700 Then x = 2700
    If x > (Me.Width - 2000) Then x = Me.Width - 2000
    LVRead.Width = x - (TV.Width + 380)
    ImgSplit.Left = x - 200
    tPesan.Left = x - 140
    
    tPesan.Width = Me.Width - (LVRead.Width + TV.Width + 240)
End Sub



Private Sub testArc_Click()

End Sub

Private Sub testArchF_Click()
    Call TestArchives
End Sub

Private Sub Tim1_Timer()
    frmProses.lblWaktu.Caption = ": " & HitungWaktu
End Sub

Private Sub TV_MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)

    With FrmUtama.TV
        If Button = vbRightButton Then
            hNode = .HitTest(x, y, False)
            If (hNode) Then
                .SelectedNode = hNode
                'Me.PopupMenu TipKhusus
            End If
        End If
    End With

End Sub

Private Sub TV_NodeClick(ByVal hNode As Long)
    MasukkanTree Alamat, hNode
    FokusNode = TV.GetNodeKey(hNode)
    FokushNode = hNode
End Sub

Private Sub UbahNama_Click()
    Call RenameArsip
End Sub

Private Sub ucTreeView1_Click()

End Sub
