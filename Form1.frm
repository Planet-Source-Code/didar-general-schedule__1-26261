VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "General Schedule"
   ClientHeight    =   4380
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   6990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4380
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6720
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2082
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":295E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":323A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3B16
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cmdlg 
      Left            =   0
      Top             =   4320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   6990
      _ExtentX        =   12330
      _ExtentY        =   1058
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Object.ToolTipText     =   "Add New Schedule"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "delete"
            Object.ToolTipText     =   "Delete Schedule"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "snd"
            Object.ToolTipText     =   "Add New Sound"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "hide"
            Object.ToolTipText     =   "Click Here To Hide"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "info"
            Object.ToolTipText     =   "Info"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Object.ToolTipText     =   "Exit To Windows System"
            ImageIndex      =   6
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   3840
         Picture         =   "Form1.frx":43F2
         ScaleHeight     =   450
         ScaleWidth      =   3000
         TabIndex        =   22
         Top             =   60
         Width           =   3000
      End
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   960
      TabIndex        =   17
      Top             =   3240
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   360
      TabIndex        =   16
      Top             =   5160
      Width           =   1095
   End
   Begin VB.ListBox List2 
      Height          =   1815
      Left            =   6840
      TabIndex        =   15
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Frame Frame3 
      Caption         =   "Date"
      Height          =   1095
      Left            =   2520
      TabIndex        =   11
      Top             =   960
      Width           =   4335
      Begin VB.CheckBox Check1 
         Caption         =   "Date"
         Height          =   255
         Left            =   3360
         TabIndex        =   14
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.Slider Slider4 
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   600
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         Max             =   31
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider Slider3 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         LargeChange     =   1
         Min             =   1
         Max             =   12
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Time"
      Height          =   1095
      Left            =   2520
      TabIndex        =   5
      Top             =   2040
      Width           =   4335
      Begin VB.Frame Frame1 
         Height          =   650
         Left            =   3360
         TabIndex        =   6
         Top             =   240
         Width           =   855
         Begin VB.OptionButton Option1 
            Caption         =   "AM"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   120
            Width           =   615
         End
         Begin VB.OptionButton Option2 
            Caption         =   "PM"
            Height          =   255
            Left            =   120
            TabIndex        =   7
            Top             =   360
            Width           =   615
         End
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         Max             =   59
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3135
         _ExtentX        =   5530
         _ExtentY        =   661
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         Max             =   12
         SelStart        =   1
         Value           =   1
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove"
      Height          =   615
      Left            =   5400
      TabIndex        =   3
      Top             =   5040
      Width           =   975
   End
   Begin VB.TextBox txtCopyText 
      Height          =   375
      Left            =   6240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   6120
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2175
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add"
      Height          =   495
      Left            =   5400
      TabIndex        =   0
      Top             =   4560
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   5880
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "© 2001Copyright By General Corporation Bangladesh. All Rights Reserved."
      Height          =   195
      Left            =   720
      TabIndex        =   21
      Top             =   4080
      Width           =   5295
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3120
      TabIndex        =   20
      Top             =   780
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter Msg"
      Height          =   195
      Left            =   120
      TabIndex        =   18
      Top             =   3240
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Label1"
      Height          =   195
      Left            =   1080
      TabIndex        =   4
      Top             =   3600
      Width           =   480
   End
   Begin VB.Menu a 
      Caption         =   "Add/delete"
      Begin VB.Menu add 
         Caption         =   "Add New Schedule"
      End
      Begin VB.Menu Delete 
         Caption         =   "Delete Schedule"
      End
   End
   Begin VB.Menu In 
      Caption         =   "Info"
      Begin VB.Menu Info 
         Caption         =   "Info"
      End
      Begin VB.Menu auto 
         Caption         =   "Set Always Start Automatically"
      End
      Begin VB.Menu never 
         Caption         =   "Never Run Automatically"
      End
      Begin VB.Menu test 
         Caption         =   "Test Sound"
      End
   End
   Begin VB.Menu Ex 
      Caption         =   "Exit"
      Begin VB.Menu about 
         Caption         =   "About"
      End
      Begin VB.Menu Exit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tnc As String
Dim sndfile As String
Dim x, X1 As Integer
      
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
 Const HKEY_LOCAL_MACHINE = &H80000002




Public Sub SaveString(hKey As Long, StrPath As String, StrValue As String, StrData As String)
   Dim KeyH&
    r = RegCreateKey(hKey, StrPath, KeyH&)
    r = RegSetValueEx(KeyH&, StrValue, 0, 1, ByVal StrData, Len(StrData))
    r = RegCloseKey(KeyH&)
End Sub






 


Private Sub about_Click()
MsgBox "General Schedule is strong schedule program.You can add thousands of schedule with different msg at a time.Put some text in the msg box and select the AM/PM." & Chr(13) & "Click 'add' to add the schedule.If you want to set the long time schedule, check the 'date' check box and select the date. And finally you have to select an audio file", 32, "About"
End Sub

Private Sub add_Click()
Command3_Click
End Sub

Private Sub auto_Click()
On Error Resume Next
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GeneralSchedule", "GeneralSchedule"
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GeneralSchedule", App.Path & "\schedule.exe"

r% = WritePrivateProfileString("sound", "status", "1", iniPath$)
If r% <> 1 Then MsgBox "An error occurred while writing SerialNumber."
     


End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Slider3.Enabled = True
Slider4.Enabled = True
Else
Slider3.Enabled = False
Slider4.Enabled = False

End If
End Sub





Private Sub Command3_Click()
Dim notes As String
Dim ast As String
Dim tahmina As Date
Dim gen As Date

On Error Resume Next

      
      
If Text1.Text = "" Then
MsgBox "You Must Enter Some MsgText", 16, "Msg"
Exit Sub
End If
      
      
      
      
If Option1.Value = True Then
ast = "AM"
Else
ast = "PM"
End If
If Option1.Value = 0 And Option2.Value = 0 Then
MsgBox "You Must Select AM/PM", 16, "AM/PM"
Exit Sub
End If



      
      
      tahmina = Slider1.Value & ":" & Slider2.Value & ":" & Second(Time) & ast
      gen = Slider3.Value & "/" & Slider4.Value
      
      
      
      List1.AddItem tahmina & "<>" & gen
      
      
For i = 0 To List1.ListCount - 1
    For x = 0 To List1.ListCount - 1
    If i = x Then GoTo Nextx
        If (List1.List(x)) = (List1.List(i)) Then
        List1.RemoveItem x
    End If
Nextx:
    Next x
Next i
txtCopyText.Text = ""

    For i = 0 To List1.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & List1.List(i) & vbCrLf
        Next i
        
FileNum = FreeFile
Open App.Path & "\history.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum


'aaaaaaaaaaaaaaaaaaaaaaaa

notes = Text1.Text

List2.AddItem notes
For i = 0 To List2.ListCount - 1
    For x = 0 To List2.ListCount - 1
    If i = x Then GoTo nexty
        If (List2.List(x)) = (List2.List(i)) Then
        List2.RemoveItem x
    End If
nexty:
    Next x
Next i
Text1.Text = ""

    For i = 0 To List2.ListCount - 1
        Text1.Text = Text1.Text & List2.List(i) & vbCrLf
        Next i
        
FileNum = FreeFile
Open App.Path & "\notes.txt" For Output As FileNum
Print #FileNum, Text1.Text
Close #FileNum


Text1.Text = ""


End Sub

Private Sub Command4_Click()
Dim a As Integer
On Error Resume Next

a = List1.ListIndex

List1.RemoveItem a
txtCopyText.Text = ""
    
    For i = 0 To List1.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & List1.List(i) & vbCrLf
        Next i
FileNum = FreeFile
Open App.Path & "\history.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum

'aaaaaaaaaaaaaaaaaaaaaaaaa

List2.RemoveItem a
txtCopyText.Text = ""
    
    For i = 0 To List2.ListCount - 1
        txtCopyText.Text = txtCopyText.Text & List2.List(i) & vbCrLf
        Next i
FileNum = FreeFile
Open App.Path & "\notes.txt" For Output As FileNum
Print #FileNum, txtCopyText.Text
Close #FileNum


End Sub



Private Sub Delete_Click()
Command4_Click
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Form_Load()
Dim a As String

 iniPath$ = App.Path + "\rwini.ini"
     sndfile = GetFromINI("sound", "file", iniPath$)
          a = GetFromINI("sound", "status", iniPath$)
     
     
     If a = "1" Then
     Me.Hide
     End If
     
     
Label3.Caption = sndfile


If Hour(Time) > 12 Then
Slider1.Value = Hour(Time) - 12
Else
Slider1.Value = Hour(Time)
End If


Slider2.Value = Minute(Time)
Slider3.Value = Month(Date)
Slider4.Value = Day(Date)

 

   
Module1.ListOpen List1, App.Path & "\history.txt"
Module1.ListOpen List2, App.Path & "\notes.txt"




 With abd
      .cbSize = Len(abd)
      .hwnd = Me.hwnd
      .uId = vbNull
      .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
      .uCallBackMessage = Mouse_Move
      .hIcon = Me.Icon
      .szTip = App.Title & ".  General Corporation Bangladesh." & vbNullChar
   End With
   Shell_NotifyIcon NIM_ADD, abd


'if no sound file is selected

If Label3.Caption = "" Then
sndfile = App.Path & "\media.wav"
Label3.Caption = sndfile
End If

 
End Sub




Private Sub Info_Click()
             MsgBox Chr(169) & " Copyright By General Corporation Bangladesh. All Rights Reserved", 32, "Info"
End Sub

Private Sub never_Click()
On Error Resume Next
SaveSetting App.Title, App.Title, "RunWithSystem", 1
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GeneralSchedule", "GeneralSchedule"
SaveSetting App.Title, App.Title, "RunWithSystem", 0
SaveString HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "GeneralSchedule", "0"



r% = WritePrivateProfileString("sound", "status", "0", iniPath$)
If r% <> 1 Then MsgBox "An error occurred while writing Status."


End Sub

Private Sub test_Click()
On Error Resume Next
PlaySound sndfile, 0&, SND_LOOP
End Sub

Private Sub Timer1_Timer()
Dim winname, pro As String
winname = Time & "<>" & Date
For X1 = 0 To List1.ListCount - 1
pro = List1.List(X1)
If winname = pro Then


           If pro = "" Then
            Exit Sub
            End If


PlaySound sndfile, 0&, SND_LOOP
MsgBox List2.List(X1), 32, "Msg"
End If
Next X1






Label1.Caption = "Current Time and Date :" & Time & "<>" & Date
End Sub









Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ItId As Long

    On Error Resume Next
    Select Case Button.Key
    
    'It's Very Important!!!!!
    'Button Name Are Case Sensitive...
    
    
        Case "add"
           Command3_Click
                    
        Case "delete"
           
         Command4_Click
            
            
        Case "snd"
        
cmdlg.Filter = "*.wav"
cmdlg.ShowOpen
sndfile = cmdlg.FileName
Label3.Caption = sndfile
     r% = WritePrivateProfileString("sound", "file", sndfile, iniPath$)
    If r% <> 1 Then MsgBox "An error occurred while writing SerialNumber."
            
            
            
            
        Case "info"
        
             MsgBox Chr(169) & " Copyright By General Corporation Bangladesh. All Rights Reserved", 32, "Info"
             
             
             
             Case "hide"
        
             Me.Hide
             
                         
            
Case "exit"

                       End
                        
                        
    End Select
    
    

End Sub






Private Sub Form_Unload(Cancel As Integer)

Shell_NotifyIcon NIM_DELETE, abd
End: End
End Sub







Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim Result As Long
Dim msg As Long
If Me.ScaleMode = vbPixels Then
     msg = x
Else
     msg = x / Screen.TwipsPerPixelX
End If

Select Case msg
             Case Mouse_Right_Down
            Case Mouse_Right_Click

Me.Show
Me.WindowState = 0


If Hour(Time) > 12 Then
Slider1.Value = Hour(Time) - 12
Else
Slider1.Value = Hour(Time)
End If


Slider2.Value = Minute(Time)
Slider3.Value = Month(Date)
Slider4.Value = Day(Date)





           
            Case Mouse_Right_DbClick
            Case Mouse_Left_Down
            Case Mouse_Left_Click
            Case Mouse_Left_DbClick
        
            Me.Show
            Me.WindowState = 0
        
        

If Hour(Time) > 12 Then
Slider1.Value = Hour(Time) - 12
Else
Slider1.Value = Hour(Time)
End If


Slider2.Value = Minute(Time)
Slider3.Value = Month(Date)
Slider4.Value = Day(Date)

        
        
        
        
        
        
            
    End Select
        

End Sub















