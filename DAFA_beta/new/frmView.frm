VERSION 5.00
Begin VB.Form frmView 
   Caption         =   "Simple View"
   ClientHeight    =   7965
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   9030
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox tView 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   9015
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
    tView.Width = frmView.Width - 255
    tView.Height = frmView.Height - 600
End Sub
