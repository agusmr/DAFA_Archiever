VERSION 5.00
Begin VB.Form frmReplace 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirm file replace"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
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
   ScaleHeight     =   3090
   ScaleWidth      =   4710
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Pic2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   150
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   15
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   1650
      Width           =   615
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   150
      ScaleHeight     =   39
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   14
      ToolTipText     =   "Pengaturan PicTmpIcon HArus seperti in ( Standarnya )"
      Top             =   810
      Width           =   615
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   285
      Left            =   3570
      TabIndex        =   13
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdRenA 
      Caption         =   "Rename all"
      Height          =   285
      Left            =   3570
      TabIndex        =   12
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   285
      Left            =   2400
      TabIndex        =   11
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdNotA 
      Caption         =   "No to all"
      Height          =   285
      Left            =   1230
      TabIndex        =   10
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdNo 
      Caption         =   "No"
      Height          =   285
      Left            =   60
      TabIndex        =   9
      Top             =   2700
      Width           =   1095
   End
   Begin VB.CommandButton cmdRen 
      Caption         =   "Rename"
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdYestA 
      Caption         =   "Yes to all"
      Height          =   285
      Left            =   1230
      TabIndex        =   7
      Top             =   2340
      Width           =   1095
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Height          =   285
      Left            =   60
      TabIndex        =   6
      Top             =   2340
      Width           =   1095
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "With this one ?"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label lUk2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   810
      TabIndex        =   4
      Top             =   1590
      Width           =   180
   End
   Begin VB.Label lUk1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   810
      TabIndex        =   3
      Top             =   810
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Would you like to replace the existing file"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   570
      Width           =   2940
   End
   Begin VB.Label lblnama 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   300
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The following file already exists"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   2250
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

