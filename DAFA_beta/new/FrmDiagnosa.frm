VERSION 5.00
Begin VB.Form FrmDiagnosa 
   Caption         =   "Darma File Archive Pesan Diagnosa Kesalahan"
   ClientHeight    =   2970
   ClientLeft      =   450
   ClientTop       =   7995
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmDiagnosa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6555
   Begin VB.CommandButton cmdClose 
      Caption         =   "Keluar"
      Height          =   435
      Left            =   4860
      TabIndex        =   1
      Top             =   2490
      Width           =   1575
   End
   Begin VB.TextBox TPesan 
      Height          =   2295
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Width           =   6285
   End
End
Attribute VB_Name = "FrmDiagnosa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    FrmUtama.Show
    Unload Me
End Sub

Private Sub Form_Load()
    KeepOnTop Me, True

End Sub
