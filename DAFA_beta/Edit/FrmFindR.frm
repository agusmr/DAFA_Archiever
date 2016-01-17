VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form FrmFindR 
   Caption         =   "Find Results"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7560
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
   ScaleHeight     =   5685
   ScaleWidth      =   7560
   StartUpPosition =   2  'CenterScreen
   Begin DAFA_Component.ucListView lvFind 
      Height          =   4155
      Left            =   60
      TabIndex        =   1
      Top             =   1020
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   7329
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      Height          =   375
      Left            =   4830
      TabIndex        =   0
      Top             =   5250
      Width           =   1695
   End
   Begin DAFA_Component.ucRebar r 
      Align           =   1  'Align Top
      Height          =   975
      Left            =   0
      Top             =   0
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   1720
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
   Begin DAFA_Component.ucToolbar t 
      Height          =   285
      Index           =   0
      Left            =   90
      Top             =   5250
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   503
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
   End
End
Attribute VB_Name = "FrmFindR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
    KeepOnTop FrmFindR, True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    lvFind.ListItems.Clear
    Cancel = True
    Me.Hide
End Sub

Private Sub t_ButtonClick(Index As Integer, ByVal Tombol As DAFA_Component.cButton)
    Select Case Tombol.Index
        Case 1
            Extract1
        Case 2
            GotoLocat CariSelect2
            Call BacaArsip
            'ViewArc CariSelect2
        Case 3
            GotoLocat CariSelect2
    End Select
End Sub

Private Function Extract1() As Long
    Dim JumSelect As Long
    
    tSimpan = TesSlash(AmbilAlamat(Alamat))
    FrmBrow.Show 1
    If tSimpan <> "" Then
        JumSelect = MasukSelect2
        If JumSelect > 0 Then
            ExtractFile Alamat, JumSelect, tSimpan
        End If
    End If

End Function
Private Function GotoLocat(ByVal Location As String) As Long
    Dim Root As String
    
    Root = TesSlash(AmbilAlamat(Location))
    TampilkanTV Root
    FrmUtama.LVRead.SetFocus
    MeSelect Location
End Function

