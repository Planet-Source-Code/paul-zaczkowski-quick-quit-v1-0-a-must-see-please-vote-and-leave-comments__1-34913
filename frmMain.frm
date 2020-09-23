VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quick Quit v1.0"
   ClientHeight    =   930
   ClientLeft      =   2490
   ClientTop       =   2175
   ClientWidth     =   4080
   ControlBox      =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   930
   ScaleWidth      =   4080
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   2760
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3240
      Top             =   240
   End
   Begin VB.PictureBox pichook 
      Height          =   555
      Left            =   180
      ScaleHeight     =   495
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   165
      Width           =   855
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   2
      Left            =   2040
      Picture         =   "frmMain.frx":0CCA
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   1
      Left            =   1560
      Picture         =   "frmMain.frx":1994
      Top             =   240
      Width           =   480
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Index           =   0
      Left            =   1080
      Picture         =   "frmMain.frx":265E
      Top             =   240
      Width           =   480
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "&Functions"
      Begin VB.Menu mnuFUN 
         Caption         =   "&Shutdown"
         Index           =   0
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "&Reboot "
         Index           =   1
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "&Logoff "
         Index           =   2
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "&Help"
         Index           =   4
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "&About"
         Index           =   5
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu mnuFUN 
         Caption         =   "&Close"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Project:   Quick Quit
'Description:  A program used to exit windows
'two mouse click, or a combination of keys.
'Combo of keys:  Ctrl+Alt+S = Shutdown
'                Ctrl+Alt+R = Reboot
'                Ctrl+Alt+L = Logoff
'Original author:  Paul Zaczkowski
'Main focus:  To teach beginners about the Windows® API.
'Date started: May 18,2002
'Date finsihed:May 18,2002
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MBUTTONDBLCLK = &H209
Private Const WM_MBUTTONDOWN = &H207
Private Const WM_MBUTTONUP = &H208
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private Const EWX_FORCE = 4
Private Const EWX_LOGOFF = 0
Private Const EWX_REBOOT = 2
Private Const EWX_SHUTDOWN = 1

Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim TrayI As NOTIFYICONDATA

Private Sub Form_Load()
If App.PrevInstance = True Then
   Unload Me
   End
End If
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd 'Link the trayicon to this picturebox
    TrayI.uId = 1&
    TrayI.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    TrayI.ucallbackMessage = WM_LBUTTONDOWN
    TrayI.hIcon = imgIcon(2).Picture
    TrayI.szTip = "Quick Quit v1.0" & Chr$(0)
    'Create the icon
    Shell_NotifyIcon NIM_ADD, TrayI
    Me.Hide
End Sub
Private Sub Form_Unload(Cancel As Integer)
    TrayI.cbSize = Len(TrayI)
    TrayI.hWnd = pichook.hWnd
    TrayI.uId = 1&
    'Delete the icon
    Shell_NotifyIcon NIM_DELETE, TrayI
    End
End Sub

Private Sub mnuFUN_Click(Index As Integer)
Dim A
Dim B
Dim C
Select Case Index
   Case 0
   A = ExitWindowsEx(EWX_SHUTDOWN, 0)
   Case 1
   B = ExitWindowsEx(EWX_REBOOT, 0)
   Case 2
   C = ExitWindowsEx(EWX_LOGOFF, 0)
   Case 5
         frmAbout.Show
   Case 7
      Unload Me
      End
   Case 4
      frmHelp.Show
   End Select
End Sub

Private Sub pichook_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then  'If the user double-clicked on the icon
        mnuFUN_Click 5
    ElseIf Msg = WM_RBUTTONUP Then  'Right click
        Me.PopupMenu mnuFunctions
    End If
End Sub
Private Sub Timer1_Timer()
    'Animate icon
    Static mPic As Integer
    Me.Icon = imgIcon(mPic).Picture
    TrayI.hIcon = imgIcon(mPic).Picture
    mPic = mPic + 1
    If mPic = 3 Then mPic = 0
    Shell_NotifyIcon NIM_MODIFY, TrayI
End Sub

Private Sub Timer2_Timer()
' Gets COMBO keys, to see if they are being inputed!
Dim CtrlKeyIsDown As Boolean
Dim AltKeyIsDown As Boolean
Dim SKeyIsDown As Boolean
Dim RKeyIsDown As Boolean
Dim KKeyIsDown As Boolean

CtrlKeyIsDown = GetAsyncKeyState(17)
AltKeyIsDown = GetAsyncKeyState(18)
SKeyIsDown = GetAsyncKeyState(83)
RKeyIsDown = GetAsyncKeyState(82)
lkeyisdown = GetAsyncKeyState(76)

If CtrlKeyIsDown And AltKeyIsDown And SKeyIsDown Then
   mnuFUN_Click 0
ElseIf CtrlKeyIsDown And AltKeyIsDown And RKeyIsDown Then
   mnuFUN_Click 1
ElseIf CtrlKeyIsDown And AltKeyIsDown And lkeyisdown Then
   mnuFUN_Click 2
Else
   Exit Sub
End If
End Sub
