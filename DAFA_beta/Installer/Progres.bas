Attribute VB_Name = "ModProg"
Public hProg1 As Long
'SetWindowLong param
Private Const GWL_STYLE As Long = (-16)
 
'Restricts input in the text box control to digits only
Private Const ES_NUMBER As Long = &H2000
Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WM_USER = &H400
Private Const PBM_SETRANGE = (WM_USER + 1)
Private Const PBM_SETPOS = (WM_USER + 2)
Private Const PBM_SETSTEP = (WM_USER + 4)
Private Const PBM_STEPIT = (WM_USER + 5)
Private Const PBM_SETRANGE32 = (WM_USER + 6)
Private Const PBM_GETRANGE = (WM_USER + 7)
Private Const PROGRESS_CLASS = "msctls_progress32"
Private Const PBS_VERTICAL = &H4
Private Const ICC_PROGRESS_CLASS = &H20

Private Declare Function CreateWindowEx Lib "user32" _
   Alias "CreateWindowExA" _
  (ByVal dwExStyle As Long, _
   ByVal lpClassName As String, _
   ByVal lpWindowName As String, _
   ByVal dwStyle As Long, _
   ByVal X As Long, ByVal Y As Long, _
   ByVal nWidth As Long, _
   ByVal nHeight As Long, _
   ByVal hWndParent As Long, _
   ByVal hMenu As Long, _
   ByVal hInstance As Long, _
   lpParam As Any) As Long


Private Declare Function SendMessage Lib "user32" _
   Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long


Public Function BuatProgres(ByVal FormNya As Form, ByVal X As Long, ByVal Y As Long, ByVal Panjang As Long, ByVal Tinggi As Long) As Long
  
   Dim hProgBar As Long
   Dim dwStyle As Long

   
   dwStyle = WS_CHILD Or WS_VISIBLE
  
   hProgBar = CreateWindowEx(0, PROGRESS_CLASS, _
                             vbNullString, _
                             dwStyle, _
                             X, Y, Panjang, Tinggi, _
                             FormNya.hwnd, 0, _
                             App.hInstance, _
                             ByVal 0)
 
    BuatProgres = hProgBar
End Function

Public Function ResetProgres(ByVal hProg As Long, ByVal Value As Long) As Long
      Call SendMessage(hProg, PBM_SETPOS, Value, ByVal 0&)
End Function



