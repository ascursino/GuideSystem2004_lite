VERSION 5.00
Begin VB.Form FrmRestIdade 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tribus Victrix (Sistema Cliente)"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmRestIdade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FrmRestIdade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   4320
      TabIndex        =   3
      Top             =   3960
      Width           =   3375
      Begin VB.CommandButton CmdEntrar 
         Caption         =   "Entrar"
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         ToolTipText     =   "Entra no sistema"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox TxtLogin 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   0
         ToolTipText     =   "Login do usuário"
         Top             =   480
         Width           =   1575
      End
      Begin VB.TextBox TxtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1200
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Senha do usuário"
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label LblLogin 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   480
         TabIndex        =   5
         Top             =   480
         Width           =   510
      End
      Begin VB.Label LblSenha 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   480
         TabIndex        =   4
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.Image ImgTribus 
      Height          =   9090
      Left            =   0
      Picture         =   "FrmRestIdade.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12090
   End
End
Attribute VB_Name = "FrmRestIdade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String

Private Sub CmdEntrar_Click()
    If TxtLogin.Text <> "" And TxtSenha <> "" Then
        
        Conecta
        
        Dim RecIdent As ADODB.Recordset
                           
        StrSql = "Select Tipo from tb_controle where Login='" & TxtLogin.Text & "' and Senha='" & TxtSenha.Text & "'"
        Set RecIdent = vgCon.Execute(StrSql)
        
        If RecIdent.EOF Then    'não achou registro
            VPStrBox = MsgBox("Login e/ou Senha não encontrados", vbInformation, "Informação")
        Else
        
            If RecIdent.Fields.Item(0).Value = "rest" Then
                VGStrRestIdade = "sim"
            ElseIf RecIdent.Fields.Item(0).Value = "naorest" Then
                VGStrRestIdade = "nao"
            ElseIf RecIdent.Fields.Item(0).Value = "sair" Then
                VGStrRestIdade = ""
            End If
            
            Unload Me
            Unload FrmTela
            
            FrmIdentifica.Show
        End If
        
        Desconecta
    Else
    
        If TxtLogin.Text = "" Then
            VPStrBox = MsgBox("Digite seu login", vbInformation, "Informação")
        ElseIf TxtSenha.Text = "" Then
            VPStrBox = MsgBox("Digite sua senha", vbInformation, "Informação")
        End If
    
    End If

End Sub

Private Sub Form_Load()
   
    Height = 9090
    Width = 12090

    Dim RetValue As Long
    RetValue = ShowWindow(Me.hwnd, SW_HIDE)

    '=== Desabilita teclas de atalho ==================
    DisableCtrlAltDelete (True)
    '================================================

    '=== Não consegue desligar o sistema via ctrl+alt+del =========
    ''Call MakeMeService
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

Private Sub TxtLogin_KeyPress(KeyAscii As Integer)
    
    '============ Símbolos ======================
    If KeyAscii >= 33 And KeyAscii <= 47 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 58 And KeyAscii <= 64 Then
        KeyAscii = 0

    ElseIf KeyAscii >= 91 And KeyAscii <= 96 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 123 And KeyAscii <= 126 Then
        KeyAscii = 0
    
    End If
    '=========================================
    
    '======= Combinações de teclas com CTRL ========
    If KeyAscii >= 1 And KeyAscii <= 7 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 9 And KeyAscii <= 29 Then
        KeyAscii = 0

    ElseIf KeyAscii = 127 Then
        KeyAscii = 0
    
    End If
    '=========================================

End Sub

Private Sub TxtSenha_KeyPress(KeyAscii As Integer)
    
    '============ Símbolos ======================
    If KeyAscii >= 33 And KeyAscii <= 47 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 58 And KeyAscii <= 64 Then
        KeyAscii = 0

    ElseIf KeyAscii >= 91 And KeyAscii <= 96 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 123 And KeyAscii <= 126 Then
        KeyAscii = 0
    
    End If
    '=========================================
    
    '======= Combinações de teclas com CTRL ========
    If KeyAscii >= 1 And KeyAscii <= 7 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 9 And KeyAscii <= 29 Then
        KeyAscii = 0

    ElseIf KeyAscii = 127 Then
        KeyAscii = 0
    
    End If
    '=========================================

End Sub
