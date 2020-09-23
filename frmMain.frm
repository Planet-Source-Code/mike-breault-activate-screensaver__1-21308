VERSION 5.00
Begin VB.Form frmScreenSaver 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activate Screen Saver"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   3750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Start Screen Saver"
      Height          =   255
      Left            =   1088
      TabIndex        =   2
      Top             =   2880
      Width           =   1575
   End
   Begin VB.PictureBox picScreen 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2535
      Left            =   495
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   2535
      ScaleWidth      =   2760
      TabIndex        =   0
      Top             =   120
      Width           =   2760
      Begin VB.PictureBox picPreview 
         Height          =   1725
         Left            =   240
         ScaleHeight     =   1665
         ScaleWidth      =   2250
         TabIndex        =   1
         Top             =   240
         Width           =   2310
      End
   End
End
Attribute VB_Name = "frmScreenSaver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Activates current screensaver
'If you use this code give me some credit

Private Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
Private Const WM_SYSCOMMAND = &H112&
Private Const SC_SCREENSAVE = &HF140&


Private Sub Command1_Click()
    Dim startSS As Long
    startSS = SendMessage(Me.hWnd, WM_SYSCOMMAND, SC_SCREENSAVE, 0&)
End Sub
