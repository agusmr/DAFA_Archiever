VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form frmOption 
   ClientHeight    =   5670
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4860
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
   ScaleHeight     =   5670
   ScaleWidth      =   4860
   StartUpPosition =   2  'CenterScreen
   Begin DAFA_Component.ucFrame fInfo 
      Height          =   4665
      Left            =   120
      Top             =   360
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   8229
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin DAFA_Component.ucProgressBar pValue 
         Height          =   3825
         Left            =   210
         Top             =   450
         Width           =   585
         _ExtentX        =   1032
         _ExtentY        =   6747
         Vertical        =   -1  'True
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total files"
         Height          =   195
         Left            =   990
         TabIndex        =   35
         Top             =   810
         Width           =   690
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         Height          =   195
         Left            =   1020
         TabIndex        =   34
         Top             =   2610
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Main comment"
         Height          =   195
         Left            =   1020
         TabIndex        =   33
         Top             =   2310
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SFX module size"
         Height          =   195
         Left            =   1020
         TabIndex        =   32
         Top             =   2010
         Width           =   1140
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Version to extract"
         Height          =   195
         Left            =   990
         TabIndex        =   31
         Top             =   420
         Width           =   1290
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ratio"
         Height          =   195
         Left            =   1020
         TabIndex        =   30
         Top             =   1620
         Width           =   375
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Packed Size"
         Height          =   195
         Left            =   1020
         TabIndex        =   29
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total size"
         Height          =   195
         Left            =   990
         TabIndex        =   28
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Archive lock"
         Height          =   195
         Left            =   1020
         TabIndex        =   27
         Top             =   3630
         Width           =   855
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Recovary record"
         Height          =   195
         Left            =   1020
         TabIndex        =   26
         Top             =   3360
         Width           =   1200
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dictionary size"
         Height          =   195
         Left            =   1020
         TabIndex        =   25
         Top             =   3090
         Width           =   1035
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Authenticity verification"
         Height          =   195
         Left            =   1020
         TabIndex        =   24
         Top             =   4080
         Width           =   1710
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   990
         X2              =   4410
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000010&
         X1              =   990
         X2              =   4410
         Y1              =   2940
         Y2              =   2940
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000010&
         X1              =   990
         X2              =   4410
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000010&
         X1              =   990
         X2              =   4410
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lVersi 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   23
         Top             =   420
         Width           =   180
      End
      Begin VB.Label lVerification 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   22
         Top             =   4080
         Width           =   180
      End
      Begin VB.Label lLock 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   21
         Top             =   3630
         Width           =   180
      End
      Begin VB.Label lRecovary 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   20
         Top             =   3360
         Width           =   180
      End
      Begin VB.Label lDictionary 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   19
         Top             =   3120
         Width           =   180
      End
      Begin VB.Label lPassword 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   18
         Top             =   2610
         Width           =   180
      End
      Begin VB.Label lCommand 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   17
         Top             =   2310
         Width           =   180
      End
      Begin VB.Label lSFX 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   16
         Top             =   2010
         Width           =   180
      End
      Begin VB.Label lRatio 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   15
         Top             =   1620
         Width           =   180
      End
      Begin VB.Label lTotPack 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   14
         Top             =   1350
         Width           =   180
      End
      Begin VB.Label lTotSize 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   13
         Top             =   1080
         Width           =   180
      End
      Begin VB.Label lTotFile 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "..."
         Height          =   195
         Left            =   4260
         TabIndex        =   12
         Top             =   810
         Width           =   180
      End
   End
   Begin DAFA_Component.ucFrame fOption 
      Height          =   4665
      Left            =   120
      Top             =   360
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   8229
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin DAFA_Component.ucFrame ucFrame8 
         Height          =   1035
         Left            =   240
         Top             =   1590
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1826
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.CheckBox ckLock 
            BackColor       =   &H80000004&
            Caption         =   "Disable archive modifications"
            Height          =   315
            Left            =   300
            TabIndex        =   10
            Top             =   360
            Width           =   2985
         End
      End
      Begin DAFA_Component.ucFrame ucFrame4 
         Height          =   1035
         Left            =   270
         Top             =   180
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1826
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.VScrollBar ScrolV 
            Enabled         =   0   'False
            Height          =   315
            Left            =   660
            TabIndex        =   8
            Top             =   540
            Width           =   285
         End
         Begin VB.TextBox tSizeRecord 
            Enabled         =   0   'False
            Height          =   345
            Left            =   300
            TabIndex        =   7
            Text            =   "0"
            Top             =   540
            Width           =   375
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Set the recovary record size to"
            Height          =   195
            Left            =   300
            TabIndex        =   9
            Top             =   240
            Width           =   2235
         End
      End
      Begin DAFA_Component.ucFrame ucFrame9 
         Height          =   1035
         Left            =   240
         Top             =   3180
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   1826
         BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
         EndProperty
         BColor          =   -2147483644
         Begin VB.CheckBox ckSign 
            BackColor       =   &H80000004&
            Caption         =   "Add authenticity information"
            Height          =   315
            Left            =   300
            TabIndex        =   11
            Top             =   390
            Width           =   2985
         End
      End
   End
   Begin DAFA_Component.ucFrame fComment 
      Height          =   4665
      Left            =   120
      Top             =   360
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   8229
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin VB.TextBox tComment 
         Height          =   3705
         Left            =   180
         MaxLength       =   30000
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   5
         Top             =   420
         Width           =   4305
      End
      Begin VB.CommandButton cmdBrow 
         Caption         =   "Load comment from file"
         Height          =   345
         Left            =   150
         TabIndex        =   4
         Top             =   4200
         Width           =   2475
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Archive comment"
         Height          =   195
         Left            =   210
         TabIndex        =   6
         Top             =   150
         Width           =   1230
      End
   End
   Begin DAFA_Component.ucFrame fSFX 
      Height          =   4665
      Left            =   120
      Top             =   360
      Width           =   4665
      _ExtentX        =   8229
      _ExtentY        =   8229
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483644
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Add SFX module"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   210
         Width           =   1155
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3510
      TabIndex        =   1
      Top             =   5130
      Width           =   1245
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   345
      Left            =   1980
      TabIndex        =   0
      Top             =   5130
      Width           =   1245
   End
   Begin DAFA_Component.ucTabStrip TbOpt 
      Height          =   5565
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   9816
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
End
Attribute VB_Name = "frmOption"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub cmdBrow_Click()
    Dim tmpPath As String
    Dim tmpHfile As Long
    Dim tmpSize As Long
    Dim tmpBit() As Byte
    Dim tmpData As String
    
    With FrmUtama.c
        If .ShowOpen Then
            tmpPath = .FileName
            tmpHfile = VbOpenFile(tmpPath)
                tmpSize = VbFileLen(tmpHfile)
                VbReadFileB tmpHfile, 1, tmpSize, tmpBit
            VbCloseHandle tmpHfile
            tmpData = StrConv(tmpBit, vbUnicode)
            tComment.Text = tmpData
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOk_Click()
    If tComment.Text <> "" Then
        Call EditCommand(Alamat, tComment.Text)
    ElseIf Header.flags And fCommand Then
        If tComment.Text <> pesan Then
            Call EditCommand(Alamat, tComment.Text)
        End If
    End If
    
    If ckLock.Enabled = True And ckLock.Value = 1 Then
        Call EditLock(Alamat)
        Call UntukLock
    End If
    
    Unload Me
End Sub
Sub buildTab()
        With TbOpt
            .Tabs.Add "Information"
            .Tabs.Add "Option"
            .Tabs.Add "Comment"
            .Tabs.Add "SFX"
        End With
        ckSign.Enabled = False
        ckLock.Value = 1
        ckLock.Enabled = False
End Sub
Sub viewInfo()
        With iArc
            lVersi.Caption = .Versi
            lTotFile.Caption = .TotFile
            lTotPack.Caption = .TotPack
            lTotSize.Caption = .TotSize
            lRatio.Caption = .Ratio
            pValue.Value = .Rat
            lSFX.Caption = .SFX
            lDictionary.Caption = .Dictionary
            lRecovary.Caption = .Recovary
            lVerification.Caption = .Verification
            lPassword.Caption = .Password
            lCommand.Caption = .Command
            lLock.Caption = .Lock
        End With
End Sub
Private Sub ucProgressBar1_DragDrop(source As Control, x As Single, y As Single)

End Sub

Private Sub Label23_Click()

End Sub

Private Sub Label22_Click()

End Sub

Private Sub fComment_Click()

End Sub

Private Sub fOption_Click()

End Sub



Private Sub Form_Load()
    Call viewInfo
    Call buildTab
    Me.Caption = FrmUtama.Caption
    
    Select Case OptFokus
        Case Is = 1
            Call TbOpt.SetSelectedTab(1)
        Case Is = 2
            Call TbOpt.SetSelectedTab(2)
        Case Is = 3
            Call TbOpt.SetSelectedTab(3)
        Case Is = 4
        
        Case Is = 5
            Call TbOpt.SetSelectedTab(2)
            ckLock.Value = 1
    End Select
End Sub

Private Sub TbOpt_Click(ByVal oTab As DAFA_Component.cTab)
    Call tampilkanOpt(oTab.Index)
End Sub

Private Sub ucFrame3_Click()

End Sub

Private Sub ucFrame6_Resize()

End Sub

