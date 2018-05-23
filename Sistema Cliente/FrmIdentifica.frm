VERSION 5.00
Begin VB.Form FrmIdentifica 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tribus Victrix (Sistema Cliente) - Identificação"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   Icon            =   "FrmIdentifica.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame FraLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Login"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   2760
      TabIndex        =   8
      Top             =   3840
      Width           =   2295
      Begin VB.TextBox TxtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   960
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Senha do usuário"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox TxtLogin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   960
         MaxLength       =   10
         TabIndex        =   0
         ToolTipText     =   "Login do usuário"
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton CmdOK 
         Caption         =   "OK"
         Height          =   255
         Left            =   840
         TabIndex        =   2
         ToolTipText     =   "Confirma dados"
         Top             =   1440
         Width           =   615
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
         Left            =   240
         TabIndex        =   10
         Top             =   840
         Width           =   555
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
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   510
      End
   End
   Begin VB.Frame FraCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cartão"
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   5760
      TabIndex        =   6
      Top             =   3840
      Width           =   3375
      Begin VB.CommandButton CmdDesistir 
         Caption         =   "Desistir"
         Height          =   255
         Left            =   1800
         TabIndex        =   5
         ToolTipText     =   "Sai do sistema"
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton CmdEntrar 
         Caption         =   "Entrar"
         Height          =   255
         Left            =   600
         TabIndex        =   4
         ToolTipText     =   "Entra no sistema"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox TxtNumCartao 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Número do cartão"
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label LblNumCartao 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nº do Cartão:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1050
      End
   End
   Begin VB.Image ImgTribus 
      Height          =   9090
      Left            =   0
      Picture         =   "FrmIdentifica.frx":000C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12090
   End
End
Attribute VB_Name = "FrmIdentifica"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrGrava As String
Public VPStrFrmCli As String

Private Sub CmdDesistir_Click()
    TxtNumCartao.Text = ""
    
    FraCartao.Enabled = False
    LblNumCartao.Enabled = False
    TxtNumCartao.Enabled = False
    CmdEntrar.Enabled = False
    'CmdSairCart.Enabled = False
    CmdDesistir.Enabled = False
    
    TxtLogin.Text = ""
    TxtSenha.Text = ""
    
    FraLogin.Enabled = True
    LblLogin.Enabled = True
    TxtLogin.Enabled = True
    LblSenha.Enabled = True
    TxtSenha.Enabled = True
    CmdOK.Enabled = True
    
End Sub

Private Sub CmdEntrar_Click()

    If TxtNumCartao.Text <> "" Then
        'conclui acesso
        
        Conecta
        
        Dim RecCart As ADODB.Recordset
        Dim RecCli As ADODB.Recordset
        Dim RecCred As ADODB.Recordset
        Dim RecConect As ADODB.Recordset
                           
        StrSql = "Select NumCartao from tb_cartao where NumCartao='" & TxtNumCartao.Text & "' and Cancelado=0"
        Set RecCart = vgCon.Execute(StrSql)

        If RecCart.EOF Then      'não achou o cartão
            VPStrBox = MsgBox("Nº do cartão não existe" & Chr(13) & "ou está cancelado.", vbInformation, "Informação")
            Desconecta
        Else        'achou cartão
            
            StrSql = "Select NumCartao,CodCli from tb_cartao where NumCartao=" & RecCart.Fields.Item(0).Value & " and CodCli=" & VGIntCodCli & ""
            Set RecCli = vgCon.Execute(StrSql)
            
            If RecCli.EOF Then  'não achou o cliente
                VPStrBox = MsgBox("Cartão não pertence ao cliente.", vbInformation, "Informação")
                       
            Else        'achou cliente
                
                StrSql = "Select TempoRest from tb_credito where NumCartao=" & RecCli.Fields.Item(0).Value
                Set RecCred = vgCon.Execute(StrSql)
                
                VGIntNumCartao = RecCli.Fields.Item(0).Value   'NumCartao
                VGStrCredito = RecCred.Fields.Item(0).Value     'TempoRest (Créditos restantes)
                VGIntMaq = Calcula_Maq(GetIPAddress())          'Pega número da máquina
                    
                 If VGStrCredito = "00:00:00" Then  'cartão não tem mais crédito
                    VPStrBox = MsgBox("Este cartão não tem mais créditos" & Chr(13) & "Para recarregá-lo, procure a recepção.", vbInformation, "Informação")
                    
                    FraLogin.Enabled = True
                    LblLogin.Enabled = True
                    TxtLogin.Enabled = True
                    TxtLogin.Text = ""
                    LblSenha.Enabled = True
                    TxtSenha.Enabled = True
                    TxtSenha.Text = ""
                    CmdOK.Enabled = True
                    
                    FraCartao.Enabled = False
                    LblNumCartao.Enabled = False
                    TxtNumCartao.Enabled = False
                    TxtNumCartao.Text = ""
                    CmdEntrar.Enabled = False
                    CmdSairCart.Enabled = False
                    
                    VGIntCodCli = 0
                    VGIntNumCartao = 0
                    VGStrCredito = ""
                    VGStrDataEntr = ""
                    VGStrDataSaida = ""
                    VGStrHoraEntr = ""
                    VGStrHoraSaida = ""
                    VGStrExecJogo = ""
                   
                 Else   'cartão tem crédito
                    StrSql = "Insert into tb_conect values (" & VGIntCodCli & "," & VGIntMaq & "," & VGIntNumCartao & ")"
                    Set RecConect = vgCon.Execute(StrSql)
                    
                    Unload Me
                    Desconecta
                    VPStrFrmCli = "existe"
                    Unload FrmCliente
                    'FrmCliente.Show
                 End If
            
            End If
            
        End If
        
    Else
        VPStrBox = MsgBox("Digite o número do cartão", vbInformation, "Informação")
    
    End If
    
    If VPStrFrmCli = "existe" Then
        VPStrFrmCli = ""
        FrmCliente.Show
    End If
    
End Sub

Private Sub CmdOK_Click()
    
    If TxtLogin.Text <> "" And TxtSenha.Text <> "" Then
        Conecta
        
        Dim RecIdent As ADODB.Recordset
        Dim RecConect  As ADODB.Recordset
        
        StrSql = "Select Tipo from tb_controle where Login='" & TxtLogin.Text & "' and Senha='" & TxtSenha.Text & "'"
        Set RecIdent = vgCon.Execute(StrSql)
        
        If Not RecIdent.EOF Then    'achou registro
        
            If RecIdent.Fields.Item(0).Value = "sair" Then
            'If (TxtLogin.Text = "sistem" Or TxtLogin.Text = "Sistem" Or TxtLogin.Text = "SISTEM") And (TxtSenha = "sistem" Or TxtSenha = "Sistem" Or TxtSenha = "SISTEM") Then
                'CmdSairCart.Enabled = True
                Call Sair
                'Desconecta
            End If
            
        Else
        
            'faz o login
            
            Dim RecLog As ADODB.Recordset
                               
            StrSql = "Select CodCli,Login,Senha from tb_acesso where Login='" & TxtLogin.Text & "' and Senha='" & TxtSenha.Text & "'"
            Set RecLog = vgCon.Execute(StrSql)
    
            If RecLog.EOF Then      'não achou o acesso
                VPStrBox = MsgBox("Login e/ou Senha não encontrados", vbInformation, "Informação")
            
            Else        'achou acesso
            
                VGIntCodCli = RecLog.Fields.Item(0).Value   'CodCli
                
                'verifica se já tem algum usuário utilizando esse login
                Call VerificaLogin
                
                If VPStrGrava = "sim" Then
                    VPStrGrava = ""
                    
                   
                    FraCartao.Enabled = True
                    LblNumCartao.Enabled = True
                    TxtNumCartao.Enabled = True
                    CmdEntrar.Enabled = True
                    CmdDesistir.Enabled = True
                    
                    FraLogin.Enabled = False
                    LblLogin.Enabled = False
                    TxtLogin.Enabled = False
                    LblSenha.Enabled = False
                    TxtSenha.Enabled = False
                    CmdOK.Enabled = False
                    'CmdSairLog.Enabled = False
                Else
                    TxtLogin.Text = ""
                    TxtSenha.Text = ""
                    TxtLogin.SetFocus
                End If
                
            End If
            
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

Sub Sair()
    
    '=== Habilita teclas de atalho ==================
    DisableCtrlAltDelete (False)
    '================================================
    
    '===== Mostra a barra de tarefas do windows ========
    Call SetWindowPos(FindWindow("Shell_TrayWnd", vbNullString), 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
    '=========================================
    
    '=========== Habilita o botão iniciar =============
    Dim RetVal As Long
    If StartButtonhWnd <> 0 Then
        RetVal = EnableWindow(StartButtonhWnd, True)
    End If
    '=========================================
    
    'Fecha o programa
    Unload Me
    Unload FrmTela
    KillProgramInMemory "Sistema Cliente.exe"
    
End Sub

Private Sub CmdSairLog_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Height = 9090
    Width = 12090
    'Top = 1590
    'Left = 1890
    
    VGIntCodCli = 0
    
    Dim RetValue As Long
    RetValue = ShowWindow(Me.hwnd, SW_HIDE)

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
    
    FraCartao.Enabled = False
    LblNumCartao.Enabled = False
    TxtNumCartao.Enabled = False
    CmdEntrar.Enabled = False
    CmdDesistir.Enabled = False
    
    VGIntCodCli = 0
    VGIntNumCartao = 0
    VGIntMaq = 0
    VGStrCredito = ""
    VGStrCredRest = ""
    VGStrDataEntr = ""
    VGStrDataSaida = ""
    VGStrHoraEntr = ""
    VGStrHoraSaida = ""
    VGStrJogo = ""
    'VGStrRestIdade = ""
    
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

Private Sub TxtNumCartao_KeyPress(KeyAscii As Integer)
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

    '============ Letras em maiúsculo e minúsculo ======================
    If KeyAscii >= 65 And KeyAscii <= 90 Then
        KeyAscii = 0
    
    ElseIf KeyAscii >= 97 And KeyAscii <= 122 Then
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

Sub VerificaLogin()
        'Conecta
        
        Dim RecVerif As ADODB.Recordset
                           
        StrSql = "Select CodCli,NumMaq from tb_conect where CodCli=" & VGIntCodCli
        Set RecVerif = vgCon.Execute(StrSql)

        If RecVerif.EOF Then    'nada foi encontrado
            VPStrGrava = "sim"
        Else
            VPStrBox = MsgBox("Esse login está sendo utilizado pelo" & Chr(13) & "cliente " & FormataNum(RecVerif.Fields.Item(0).Value) & " na máquina " & FormataNum(RecVerif.Fields.Item(1).Value) & ".", vbInformation, "Informação")
        End If
        
        'Desconecta
End Sub
