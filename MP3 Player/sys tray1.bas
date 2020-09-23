Attribute VB_Name = "Module1"
Option Explicit
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const LB_SETHORIZONTALEXTENT = &H194
      Type NOTIFYICONDATA
         cbSize As Long
         hwnd As Long
         uId As Long
         uFlags As Long
         uCallBackMessage As Long
         hIcon As Long
         szTip As String * 64
      End Type


      Global Const NIM_ADD = &H0
      Global Const NIM_MODIFY = &H1
      Global Const NIM_DELETE = &H2

      Global Const WM_MOUSEMOVE = &H200
      
      Global Const NIF_MESSAGE = &H1
      Global Const NIF_ICON = &H2
      Global Const NIF_TIP = &H4


      Global Const WM_LBUTTONDBLCLK = &H203
      Global Const WM_LBUTTONDOWN = &H201
      Global Const WM_LBUTTONUP = &H202

      
      Global Const WM_RBUTTONDBLCLK = &H206
      Global Const WM_RBUTTONDOWN = &H204
      Global Const WM_RBUTTONUP = &H205

      Declare Function Shell_NotifyIcon Lib "shell32" _
         Alias "Shell_NotifyIconA" _
         (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean


      Global nid As NOTIFYICONDATA
Sub AddToTray(TrayIcon, TrayText As String, TrayForm As Form)
       
         nid.cbSize = Len(nid)
         nid.hwnd = TrayForm.hwnd
         nid.uId = vbNull
         nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
         nid.uCallBackMessage = WM_MOUSEMOVE
         nid.hIcon = TrayIcon
         nid.szTip = TrayText & vbNullChar

         Shell_NotifyIcon NIM_ADD, nid
         TrayForm.Hide
End Sub
Function RespondToTray(X As Single)
          RespondToTray = 0
          Dim msg As Long
          Dim sFilter As String
          If Form1.ScaleMode <> 3 Then msg = X / Screen.TwipsPerPixelX Else: msg = X
          Select Case msg
             Case WM_LBUTTONDOWN
             RespondToTray = 2
             Case WM_RBUTTONDOWN
             RespondToTray = 2
             Case WM_RBUTTONUP
             Case WM_RBUTTONDBLCLK
          End Select
End Function
      Sub ShowFormAgain(TrayForm As Form)
      Call RemoveFromTray
      TrayForm.Show
      End Sub
Sub RemoveFromTray()
Shell_NotifyIcon NIM_DELETE, nid
End Sub
