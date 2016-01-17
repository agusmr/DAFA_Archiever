VERSION 5.00
Begin VB.Form FrmUninstall 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Uninstall DAFA"
   ClientHeight    =   2625
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   3975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   3975
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMulai 
      Caption         =   "Uninstall"
      Height          =   885
      Left            =   570
      TabIndex        =   0
      Top             =   900
      Width           =   2625
   End
End
Attribute VB_Name = "FrmUninstall"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdMulai_Click()
    Call Uninstall
    MsgBox "Uninstall Complite !", vbInformation
End Sub
Private Sub Uninstall()
    DeleteKey "HKCR\.gus"
    DeleteKey "HKCR\*\shell\Add to Gus Archive Her\command\"
    DeleteKey "HKCR\*\shell\Add to Gus Archive Her\"
    DeleteKey "HKCR\Gus File\DefaultIcon\"
    DeleteKey "HKCR\Gus File\shell\Extract With Gus\command\"
    DeleteKey "HKCR\Gus File\shell\Extract With Gus\"
    DeleteKey "HKCR\Gus File\shell\open\command\"
    DeleteKey "HKCR\Gus File\shell\open\"
    DeleteKey "HKCR\Gus File\shell\"
    DeleteKey "HKCR\Gus File\"
    DeleteKey "HKLM\SOFTWARE\Classes\Folder\shell\Add to Gus Archive Her\command\"
    DeleteKey "HKLM\SOFTWARE\Classes\Folder\shell\Add to Gus Archive Her\"
    DeleteKey "HKLM\SOFTWARE\Classes\Folder\shell\Add to Gus Archive..\command\"
    DeleteKey "HKLM\SOFTWARE\Classes\Folder\shell\Add to Gus Archive..\"
    DeleteKey "HKCR\WinRAR\shell\Extract with Gus\command\"
    DeleteKey "HKCR\WinRAR\shell\Extract with Gus\"
    DeleteKey "HKCR\Rax File\shell\Extract with Gus\command\"
    DeleteKey "HKCR\Rax File\shell\Extract with Gus\"
    DeleteKey "HKCR\WinRAR\shell\Open With Gus\command\"
    DeleteKey "HKCR\WinRAR\shell\Open With Gus\"
    DeleteKey "HKCR\RaX File\shell\Open With Gus\command\"
    DeleteKey "HKCR\RaX File\shell\Open With Gus\"
End Sub
Public Sub DeleteKey(Value As String)

    Dim b As Object
    On Error Resume Next
    Set b = CreateObject("Wscript.Shell")
    b.RegDelete Value

End Sub


