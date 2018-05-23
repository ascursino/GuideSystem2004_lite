VERSION 5.00
Begin VB.Form FrmTela 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tribus Victrix (Sistema Cliente)"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmTela.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   240
      Top             =   120
   End
   Begin VB.Image ImgTribus 
      Height          =   9090
      Left            =   0
      Picture         =   "FrmTela.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12090
   End
End
Attribute VB_Name = "FrmTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Height = 9090
    Width = 12090
    
    '=== Desabilita teclas de atalho ==================
    DisableCtrlAltDelete (True)
    '================================================

    '=== Não consegue desligar o sistema via ctrl+alt+del =========
    Call MakeMeService
    '================================================

    '====== Oculta a barra de tarefas do windows ===============
    Call SetWindowPos(FindWindow("Shell_TrayWnd", vbNullString), 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    '================================================

    '======= Desabilita o botão iniciar ======================
    Dim WinClass As String
    Dim TaskBarHwnd As Long, lRet As Long, lParam As Long
    '
    WinClass = "Shell_TrayWnd"
    '
    TaskBarHwnd = FindWindow(WinClass, vbNullString)
    '
    lRet = EnumChildWindows(TaskBarHwnd, AddressOf EnumChildProc, lParam)
    '================================================

    '================ Prende mouse no form ===============
    'Dim EstaJanela As Retang

    'GetWindowRect Me.hWnd, EstaJanela
    'ClipCursor EstaJanela
    '================================================

End Sub

Private Sub Timer1_Timer()
    DisableCtrlAltDelete (True)
End Sub
