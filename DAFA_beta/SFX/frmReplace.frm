VERSION 5.00
Begin VB.Form frmReplace 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdYestA 
      Caption         =   "Yes to all"
      Height          =   285
      Left            =   1170
      TabIndex        =   8
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdRen 
      Caption         =   "Rename"
      Height          =   285
      Left            =   2340
      TabIndex        =   7
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   285
      Left            =   0
      TabIndex        =   6
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdNotA 
      Caption         =   "No to all"
      Height          =   285
      Left            =   1170
      TabIndex        =   5
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   2340
      TabIndex        =   4
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdRenA 
      Caption         =   "Rename all"
      Height          =   285
      Left            =   3510
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   285
      Left            =   3510
      TabIndex        =   2
      Top             =   2640
      Width           =   1095
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   90
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   1
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   750
      Width           =   615
   End
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   90
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   0
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   1590
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The following file already exists"
      Height          =   195
      Left            =   60
      TabIndex        =   15
      Top             =   0
      Width           =   2250
   End
   Begin VB.Label lblnama 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   60
      TabIndex        =   14
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to replace the existing file"
      Height          =   195
      Left            =   60
      TabIndex        =   13
      Top             =   510
      Width           =   2940
   End
   Begin VB.Label lUk1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   750
      TabIndex        =   12
      Top             =   750
      Width           =   180
   End
   Begin VB.Label lUk2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   750
      TabIndex        =   11
      Top             =   1530
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "With this one ?"
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   1380
      Width           =   1065
   End
End
Attribute VB_Name = "frmReplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Yes As Boolean

Private Sub cmdCancel_Click()
    Yes = False
    NotoAll = True
    Unload Me
End Sub

Private Sub cmdNo_Click()
    Yes = False
    Unload Me
End Sub

Private Sub cmdNotA_Click()
    Yes = False
    NotoAll = True
    Unload Me
End Sub

Private Sub cmdRen_Click()
    Dim Name As String
    
    Yes = True
    Name = NameRename
    NameRename = InputBox("Masukkan nama file baru", "Darma file archive", AmbilNama(NameRename))
    NameRename = TesSlash(AmbilAlamat(Name)) & NameRename
    Unload Me
End Sub

Private Sub cmdRenA_Click()
    Yes = True
    NameRename = NameRename
    RenameAll = True
    Unload Me
End Sub

Private Sub cmdYes_Click()
    Yes = True
    NameRename = NameRename
    Unload Me
End Sub

Private Sub cmdYestA_Click()
    Yes = True
    YestoAll = True
    NameRename = NameRename
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Yes = False Then NameRename = ""
    Unload Me
End Sub


