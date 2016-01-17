VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form FrmFind 
   Caption         =   "Find files"
   ClientHeight    =   3015
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5835
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
   ScaleHeight     =   3015
   ScaleWidth      =   5835
   StartUpPosition =   2  'CenterScreen
   Begin DAFA_Component.ucFrame ucFrame1 
      Height          =   2025
      Left            =   180
      Top             =   840
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   3572
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      Begin VB.ComboBox cString 
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   1200
         Width           =   2985
      End
      Begin VB.ComboBox cName 
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   480
         Width           =   2985
      End
      Begin VB.CheckBox ckCase 
         Caption         =   "Match case"
         Enabled         =   0   'False
         Height          =   225
         Left            =   90
         TabIndex        =   5
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "String to find"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   930
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "File names to find"
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   180
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3990
      TabIndex        =   3
      Top             =   930
      Width           =   1695
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4020
      TabIndex        =   2
      Top             =   2490
      Width           =   1695
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   3990
      TabIndex        =   1
      Top             =   150
      Width           =   1695
   End
   Begin VB.ComboBox cAlamat 
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   450
      Width           =   3195
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Name archive to find"
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   150
      Width           =   1485
   End
End
Attribute VB_Name = "FrmFind"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    If cName.text <> "" Then
        FindArchive Alamat, cName.text
        Unload Me
    Else
        MsgBox "Masukkan kata kunci !", vbCritical
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Me.Hide
End Sub
