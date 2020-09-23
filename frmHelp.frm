VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next >>"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3600
      Width           =   4335
   End
   Begin VB.Label Label2 
      BackColor       =   &H0000FFFF&
      Caption         =   $"frmHelp.frx":0CCA
      Height          =   1815
      Left            =   2640
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Right click on the icon (shown below) to open the  menu."
      Height          =   1455
      Left            =   2760
      TabIndex        =   2
      Top             =   1200
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   2175
      Left            =   1080
      Picture         =   "frmHelp.frx":0D56
      Top             =   1080
      Width           =   2580
   End
   Begin VB.Label lbl1 
      BackStyle       =   0  'Transparent
      Caption         =   "1."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label lblIntro 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "So, you're having some trouble on how to work this thing.  Well just follow the steps below, and you'll be fixed up in no time!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   0
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   825
      Left            =   0
      Picture         =   "frmHelp.frx":131DC
      Stretch         =   -1  'True
      Top             =   0
      Width           =   1440
   End
   Begin VB.Image Image3 
      Height          =   2415
      Left            =   840
      Picture         =   "frmHelp.frx":4321E
      Stretch         =   -1  'True
      Top             =   4320
      Width           =   3060
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdNext_Click()
   Image2.Picture = Image3.Picture
   Label2.Left = 2640
   Label2.Top = 1080
   cmdNext.Visible = False
   cmdExit.Visible = True
   lbl1.Caption = "2.  Click to..."
End Sub

