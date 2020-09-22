Attribute VB_Name = "Module2"

Option Explicit
'Read-Write INI Sample
'Written by: George Csefai-Keane, Inc.
'email: george.csefai@keaneinc.com

'API Declarations
 Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long






 Declare Function PlaySound Lib "Winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
     Declare Function Shell_NotifyIcon Lib "SHELL32" _
      Alias "Shell_NotifyIconA" _
      (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
      Public abd As NOTIFYICONDATA

      Public Type NOTIFYICONDATA
       cbSize As Long
       hwnd As Long
       uId As Long
       uFlags As Long
       uCallBackMessage As Long
       hIcon As Long
       szTip As String * 64
      End Type

      Public Const NIF_MESSAGE = &H1
      Public Const NIF_ICON = &H2
      Public Const NIF_TIP = &H4
      Public Const Mouse_Move = 512
      Public Const Mouse_Left_Down = 513
      Public Const Mouse_Left_Click = 514
      Public Const Mouse_Left_DbClick = 515
      Public Const Mouse_Right_Down = 516
      Public Const Mouse_Right_Click = 517
      Public Const Mouse_Right_DbClick = 518
      Public Const Mouse_Button_Down = 519
      Public Const Mouse_Button_Click = 520
      Public Const Mouse_Button_DbClick = 521















Global r%
Global entry$
Global iniPath$

Sub CenterForm(frm As Form)
    frm.Top = (Screen.Height * 0.85) / 2 - frm.Height / 2
    frm.Left = Screen.Width / 2 - frm.Width / 2
End Sub

Function GetFromINI(AppName$, KeyName$, FileName$) As String
   Dim RetStr As String
   RetStr = String(255, Chr(0))
   GetFromINI = Left(RetStr, GetPrivateProfileString(AppName$, ByVal KeyName$, "", RetStr, Len(RetStr), FileName$))
End Function



