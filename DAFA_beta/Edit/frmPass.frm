VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masukkan Password"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   222
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPass.frx":0000
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   3825
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   375
      Left            =   2130
      TabIndex        =   3
      Top             =   930
      Width           =   1365
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   930
      Width           =   1365
   End
   Begin VB.TextBox TPassword 
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   180
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   300
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Masukkan Password :"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   60
      Width           =   1545
   End
End
Attribute VB_Name = "frmPass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdKeluar_Click()
    If Terbuka Then
        Unload Me
    Else
        End
    End If
End Sub

Private Sub cmdOk_Click()
    pass = TPassword.Text
    Unload Me
    
End Sub
Private Sub Form_Load()
    KeepOnTop Me, True
    TPassword.Text = ""
End Sub

Private Sub TPassword_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdOk_Click

End Sub
