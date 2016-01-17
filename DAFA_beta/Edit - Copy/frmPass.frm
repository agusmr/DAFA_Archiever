VERSION 5.00
Begin VB.Form frmPass 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masukkan Password"
   ClientHeight    =   1815
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   4650
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
   ScaleHeight     =   1815
   ScaleWidth      =   4650
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdKeluar 
      Caption         =   "Keluar"
      Height          =   495
      Left            =   2670
      TabIndex        =   2
      Top             =   1080
      Width           =   1605
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   810
      TabIndex        =   1
      Top             =   1080
      Width           =   1605
   End
   Begin VB.TextBox TPassword 
      Height          =   495
      IMEMode         =   3  'DISABLE
      Left            =   840
      PasswordChar    =   "*"
      TabIndex        =   0
      Top             =   240
      Width           =   3465
   End
   Begin VB.Image Image1 
      Height          =   840
      Left            =   0
      Picture         =   "frmPass.frx":030A
      Stretch         =   -1  'True
      Top             =   120
      Width           =   840
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
