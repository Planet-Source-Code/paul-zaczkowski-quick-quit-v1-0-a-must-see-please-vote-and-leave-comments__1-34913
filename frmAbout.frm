VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About Quick Quit"
   ClientHeight    =   3555
   ClientLeft      =   4965
   ClientTop       =   3075
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   1560
      Top             =   1080
   End
   Begin VB.TextBox txtDescription 
      Height          =   1095
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmAbout.frx":0CCA
      Top             =   1200
      Width           =   3975
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":0E8C
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Top             =   2625
      Width           =   1260
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   240
      Picture         =   "frmAbout.frx":1ACE
      Top             =   1800
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   240
      Picture         =   "frmAbout.frx":2710
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmAbout.frx":3352
      Top             =   840
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000001&
      X1              =   0
      X2              =   5760
      Y1              =   2490
      Y2              =   2490
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000014&
      X1              =   0
      X2              =   5760
      Y1              =   2505
      Y2              =   2505
   End
   Begin VB.Label lblTitle 
      Caption         =   "Quick Quit "
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version:  1.0.0"
      Height          =   225
      Left            =   1080
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"frmAbout.frx":3F94
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   2
      Top             =   2640
      Width           =   3870
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim A As Integer

Private Sub cmdOK_Click()
   Me.Hide
End Sub

Private Sub Form_Load()
   txtDescription.Text = "Quick Quit, is not in any way licensed or endorsed by Microsoft®.  You may freely copy and distribute this product as you please.  It is freeware, so the author takes no credit for any damages that this product may inflict on your computer.  Any abused used of this product, by the user, cannot be fined, but, the author takes no responsibilty for the users actions." & vbNewLine & vbNewLine & "(Ctrl+Alt+R>>Reboot;  Ctrl+Alt+S>>Shutdown; Ctrl+Alt+L>>Logoff!)"

';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;'
''Quick Quit, is not in any way licensed or endorsed
''by Microsoft®.  You may freely copy and distribute
''this product as you please.  It is freeware, so the
''author takes no credit for any damages that this
''product may inflict on your computer.
''Any abused used of this product, by the user,
''cannot be fined, but, the author takes no
''responsibilty for the users actions.
';;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;'
A = 0
End Sub

Private Sub Timer1_Timer()
   If A = 3 Then
      A = 0
   End If
   picIcon.Picture = Image1(A).Picture
   A = A + 1
End Sub
