VERSION 5.00
Object = "{5A698B5B-F00C-49C2-9AFD-AA4A490CF2B8}#1.0#0"; "Componen.ocx"
Begin VB.Form FrmSetting 
   BackColor       =   &H8000000E&
   Caption         =   "Setting"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   7545
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
   ScaleHeight     =   4365
   ScaleWidth      =   7545
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   435
      Left            =   5760
      TabIndex        =   32
      Top             =   3870
      Width           =   1725
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   435
      Left            =   3780
      TabIndex        =   31
      Top             =   3870
      Width           =   1725
   End
   Begin VB.TextBox tTemp 
      Enabled         =   0   'False
      Height          =   435
      Left            =   3240
      TabIndex        =   30
      Top             =   3270
      Width           =   4245
   End
   Begin DAFA_Component.ucFrame ucFrame2 
      Height          =   2295
      Left            =   3240
      Top             =   30
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   4048
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483634
      Begin VB.CheckBox cToolbar 
         BackColor       =   &H8000000E&
         Caption         =   "Show buttons text"
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1020
         Width           =   1875
      End
      Begin VB.CheckBox cToolbar 
         BackColor       =   &H8000000E&
         Caption         =   "Large buttons"
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   120
         TabIndex        =   34
         Top             =   1290
         Width           =   1875
      End
      Begin VB.CheckBox cToolbar 
         BackColor       =   &H8000000E&
         Caption         =   "Lock toolbar"
         Height          =   315
         Index           =   5
         Left            =   120
         TabIndex        =   33
         Top             =   1590
         Width           =   1875
      End
      Begin VB.CheckBox cToolbar 
         BackColor       =   &H8000000E&
         Caption         =   "View ""Up one level"" buttons"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   450
         Width           =   2295
      End
      Begin VB.CheckBox cToolbar 
         BackColor       =   &H8000000E&
         Caption         =   "View main toolbar"
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1875
      End
      Begin VB.CheckBox cToolbar 
         BackColor       =   &H8000000E&
         Caption         =   "View address bar"
         Height          =   315
         Index           =   2
         Left            =   120
         TabIndex        =   0
         Top             =   720
         Width           =   1875
      End
   End
   Begin DAFA_Component.ucFrame ucFrame3 
      Height          =   3165
      Left            =   60
      Top             =   60
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   5583
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483634
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "New"
         Enabled         =   0   'False
         Height          =   315
         Index           =   19
         Left            =   1680
         TabIndex        =   36
         Top             =   2760
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Repair"
         Enabled         =   0   'False
         Height          =   315
         Index           =   11
         Left            =   1680
         TabIndex        =   21
         Top             =   420
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "View"
         Height          =   315
         Index           =   3
         Left            =   300
         TabIndex        =   20
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Delete"
         Height          =   315
         Index           =   4
         Left            =   300
         TabIndex        =   19
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Find"
         Height          =   315
         Index           =   5
         Left            =   300
         TabIndex        =   18
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Print"
         Enabled         =   0   'False
         Height          =   315
         Index           =   6
         Left            =   300
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Wizard"
         Enabled         =   0   'False
         Height          =   315
         Index           =   7
         Left            =   300
         TabIndex        =   16
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Extract"
         Height          =   315
         Index           =   12
         Left            =   1680
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Comment"
         Height          =   315
         Index           =   14
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Protect"
         Enabled         =   0   'False
         Height          =   315
         Index           =   15
         Left            =   1680
         TabIndex        =   13
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Lock"
         Height          =   315
         Index           =   16
         Left            =   1680
         TabIndex        =   12
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "SFX"
         Enabled         =   0   'False
         Height          =   315
         Index           =   17
         Left            =   1680
         TabIndex        =   11
         Top             =   2220
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Exit"
         Height          =   315
         Index           =   10
         Left            =   1680
         TabIndex        =   10
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Report"
         Enabled         =   0   'False
         Height          =   315
         Index           =   18
         Left            =   1680
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Convert"
         Enabled         =   0   'False
         Height          =   315
         Index           =   8
         Left            =   300
         TabIndex        =   8
         Top             =   2520
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Add"
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Info"
         Height          =   315
         Index           =   9
         Left            =   300
         TabIndex        =   6
         Top             =   2820
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Test"
         Height          =   315
         Index           =   2
         Left            =   300
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "VirusScan"
         Enabled         =   0   'False
         Height          =   315
         Index           =   13
         Left            =   1680
         TabIndex        =   4
         Top             =   1020
         Width           =   1095
      End
      Begin VB.CheckBox cButton 
         BackColor       =   &H8000000E&
         Caption         =   "Extract to"
         Height          =   315
         Index           =   1
         Left            =   300
         TabIndex        =   3
         Top             =   420
         Width           =   1095
      End
   End
   Begin DAFA_Component.ucFrame ucFrame4 
      Height          =   855
      Left            =   3240
      Top             =   2370
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   1508
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483634
      Begin VB.CheckBox cShowGrid 
         BackColor       =   &H8000000E&
         Caption         =   "Show grid lines"
         Height          =   315
         Left            =   150
         TabIndex        =   23
         Top             =   150
         Width           =   1575
      End
      Begin VB.CheckBox cFullselect 
         BackColor       =   &H8000000E&
         Caption         =   "Full row select"
         Height          =   315
         Left            =   150
         TabIndex        =   22
         Top             =   450
         Width           =   1575
      End
   End
   Begin DAFA_Component.ucFrame ucFrame5 
      Height          =   3195
      Left            =   5790
      Top             =   30
      Width           =   1725
      _ExtentX        =   3043
      _ExtentY        =   5636
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483634
      Begin VB.OptionButton OptView 
         BackColor       =   &H8000000E&
         Caption         =   "Detail"
         Height          =   255
         Index           =   1
         Left            =   270
         TabIndex        =   28
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton OptView 
         BackColor       =   &H8000000E&
         Caption         =   "Small icons"
         Height          =   255
         Index           =   2
         Left            =   270
         TabIndex        =   27
         Top             =   810
         Width           =   1215
      End
      Begin VB.OptionButton OptView 
         BackColor       =   &H8000000E&
         Caption         =   "Icons"
         Height          =   255
         Index           =   0
         Left            =   270
         TabIndex        =   26
         Top             =   150
         Width           =   1215
      End
      Begin VB.OptionButton OptView 
         BackColor       =   &H8000000E&
         Caption         =   "List"
         Height          =   255
         Index           =   3
         Left            =   270
         TabIndex        =   25
         Top             =   1140
         Width           =   1215
      End
      Begin VB.OptionButton OptView 
         BackColor       =   &H8000000E&
         Caption         =   "Title"
         Height          =   255
         Index           =   4
         Left            =   270
         TabIndex        =   24
         Top             =   1470
         Width           =   1215
      End
   End
   Begin DAFA_Component.ucFrame ucFrame6 
      Height          =   585
      Left            =   60
      Top             =   3150
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   1032
      BeginProperty Font {8EE14374-8533-49EE-AF89-86E60079ACE7} 
      EndProperty
      BColor          =   -2147483634
      Begin VB.CheckBox cLog 
         BackColor       =   &H8000000E&
         Caption         =   "Create log"
         Enabled         =   0   'False
         Height          =   315
         Left            =   150
         TabIndex        =   29
         Top             =   150
         Width           =   1575
      End
   End
End
Attribute VB_Name = "FrmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Call CreateSetting
    Call LoadSetting
    Unload Me
End Sub

Private Sub Form_Load()
    Call LoadSetting
End Sub
