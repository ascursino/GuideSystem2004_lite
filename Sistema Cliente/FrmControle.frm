VERSION 5.00
Begin VB.Form FrmControle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tribus Victrix (Sistema Cliente) - Controle"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4800
   ControlBox      =   0   'False
   Icon            =   "FrmControle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   4320
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   4320
      Top             =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4320
      Top             =   120
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   255
      Left            =   4080
      TabIndex        =   11
      ToolTipText     =   "Sai do sistema"
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label LblMaq 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Máquina:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   2280
      TabIndex        =   10
      Top             =   1440
      Width           =   795
   End
   Begin VB.Label LblResMaq 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label LblDataHoraAtual 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   210
      Left            =   2205
      TabIndex        =   8
      Top             =   720
      Width           =   645
   End
   Begin VB.Label LblResDataEntr 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00/00/0000  00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   210
      Left            =   1680
      TabIndex        =   7
      Top             =   2160
      Width           =   1530
   End
   Begin VB.Label LblResCredRest 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00:00:00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   285
      Left            =   960
      TabIndex        =   6
      Top             =   1800
      Width           =   990
   End
   Begin VB.Label LblResCartao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   210
      Left            =   840
      TabIndex        =   5
      Top             =   1440
      Width           =   180
   End
   Begin VB.Label LblResCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(nome cliente)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   210
      Left            =   840
      TabIndex        =   4
      Top             =   1200
      Width           =   1020
   End
   Begin VB.Label LblDataEntr 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data/Hora entrada:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   2160
      Width           =   1485
   End
   Begin VB.Label LblCredRest 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Créditos:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   765
   End
   Begin VB.Label LblCartao 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cartão:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   585
   End
   Begin VB.Label LblCliente 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   630
   End
   Begin VB.Image ImgTribus 
      Height          =   495
      Left            =   720
      Picture         =   "FrmControle.frx":000C
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "FrmControle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String

Private Sub CmdSair_Click()
    
    'grava no banco informações de tempo restante
    
    Dim RecCred As ADODB.Recordset
    Dim RecMaq As ADODB.Recordset
    Dim RecMaqCli As ADODB.Recordset
    Dim RecConect As ADODB.Recordset
    Dim RET As Long
    
    Conecta
    
    StrSql = "Update tb_credito set TempoRest='" & restante & "' where NumCartao=" & VGIntNumCartao
    Set RecCred = vgCon.Execute(StrSql)
    
    StrSql = "Update tb_maquina set Situacao='livre' where NumMaq=" & VGIntMaq
    Set RecMaq = vgCon.Execute(StrSql)
    
    StrSql = "Update tb_maqcli set DataSaida='" & FormataDataUS(Date) & "',HoraSaida='" & VPStrHoraSaida & "' where CodCli=" & VGIntCodCli
    Set RecMaqCli = vgCon.Execute(StrSql)
    
    StrSql = "Delete from tb_conect where CodCli=" & VGIntCodCli
    Set RecConect = vgCon.Execute(StrSql)
    
    Desconecta
    
    VGIntIdade = 0
    VGIntCodCli = 0
    VGIntNumCartao = 0
    VGStrCredito = ""
    VGIntMaq = 0
    
    hora_inicial = 0
    hora_atual = 0
    qtde = 0
    tempo_user = 0
    
    Unload Me
    Unload FrmCliente
    Unload FrmControle
    'Unload FrmTela
    
    '========== Faz logoff do usuário =================
    'RET = ExitWindowsEx(EWX_LOGOFF, 0)
    '===========================================
    
    '========== Desligar o computador ================
    'RET = ExitWindowsEx(EWX_SHUTDOWN, 0)
    '===========================================
    
    '========== Reiniciar o computador ===============
    'RET = ExitWindowsEx(EWX_REBOOT, 0)
    '===========================================
    
    'VPStrBox = MsgBox("Seu crédito restante é de " & VGStrCredRest, vbInformation, "Informação")
    restante = 0
    Unload FrmTela
    FrmIdentifica.Show
End Sub

Private Sub Form_Load()
    Height = 2805
    Width = 4890
    Top = 2880
    Left = 3555
    
    CmdSair.Visible = False
    
    FrmTela.Enabled = False
    
    '=== Desabilita teclas de atalho ==================
    DisableCtrlAltDelete (True)
    '================================================

    '===== Não consegue desligar o sistema via ctrl+alt+del ========
    Call MakeMeService
    '=================================================

    '====== Oculta a barra de tarefas do windows ===============
    Call SetWindowPos(FindWindow("Shell_TrayWnd", vbNullString), 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
    '================================================

    '======= Desabilita o botão iniciar ===================
    Dim WinClass As String
    Dim TaskBarHwnd As Long, lRet As Long, lParam As Long

    WinClass = "Shell_TrayWnd"

    TaskBarHwnd = FindWindow(WinClass, vbNullString)

    lRet = EnumChildWindows(TaskBarHwnd, AddressOf EnumChildProc, lParam)
    '=============================================

    Timer1.Enabled = True
    Timer1.Interval = 1000

    hora_inicial = Now
    tempo_user = VGStrCredito   'tempo de crédito
    LblResCredRest.Caption = tempo_user

    LblResCliente.Caption = FrmCliente.LblResNome
    LblResCartao.Caption = FormataNum(VGIntNumCartao)
    LblResMaq.Caption = FormataNum(VGIntMaq)
    LblResDataEntr = hora_inicial

    Dim i&

    If VGStrJogo = "strike" Then
        VGStrJogo = ""

        'inicia jogo Counter Strike
        'i& = ShellExecute(0, "open", "C:\windows\sol.exe", "", "C:\windows", SW_SHOW)
        i& = ShellExecute(0, "open", "C:\Sierra\Half-Life\hl.exe", "-console -game cstrike", "C:\Sierra\Half-Life", SW_SHOW)
        'i& = ShellExecute(0, "open", "C:\Sierra\Half-Life\cstrike\autoexec.bat", "", "C:\Sierra\Half-Life\cstrike", SW_SHOW)
        VGStrExecJogo = "hl.exe"

    ElseIf VGStrJogo = "half" Then
        VGStrJogo = ""

        'inicia jogo Half Life
        i& = ShellExecute(0, "open", "C:\Sierra\Half-Life\hl.exe", "", "C:\Sierra\Half-Life", SW_SHOW)
        VGStrExecJogo = "hl.exe"

    ElseIf VGStrJogo = "war" Then
        VGStrJogo = ""

        'inicia jogo WarCraft III
        i& = ShellExecute(0, "open", "C:\Arquivos de programas\Warcraft III\Warcraft III.exe", "", "C:\Arquivos de programas\Warcraft III", SW_SHOW)
        VGStrExecJogo = "Warcraft III.exe"
        VGStrExecJogo1 = "War3.exe"

    ElseIf VGStrJogo = "death" Then
        VGStrJogo = ""

        'inicia jogo DeathMatch Classic
        i& = ShellExecute(0, "open", "C:\Sierra\Half-Life\hl.exe", "-game dmc", "C:\Sierra\Half-Life", SW_SHOW)
        VGStrExecJogo = "hl.exe"

    ElseIf VGStrJogo = "day" Then
        VGStrJogo = ""

        'inicia jogo Day of Defeat
        i& = ShellExecute(0, "open", "C:\Sierra\Half-Life\hl.exe", "-console -game dod", "C:\Sierra\Half-Life", SW_SHOW)
        VGStrExecJogo = "hl.exe"

    ElseIf VGStrJogo = "bat" Then
        VGStrJogo = ""

        'inicia jogo Battle Field 1942
        i& = ShellExecute(0, "open", "C:\Arquivos de programas\EA GAMES\Battlefield 1942\BF1942.exe", "", "C:\Arquivos de programas\EA GAMES\Battlefield 1942", SW_SHOW)
        VGStrExecJogo = "BF1942.exe"

    ElseIf VGStrJogo = "age" Then
        VGStrJogo = ""

        'inicia jogo Age of Empires
        i& = ShellExecute(0, "open", "C:\Arquivos de programas\Microsoft Games\Age of Empires II\empires2.exe", "", "C:\Arquivos de programas\Microsoft Games\Age of Empires II", SW_SHOW)
        VGStrExecJogo = "empires2.exe"

    ElseIf VGStrJogo = "agemit" Then
        VGStrJogo = ""

        'inicia jogo Age of Mythology
        i& = ShellExecute(0, "open", "C:\Arquivos de programas\Microsoft Games\Age of Mythology\aom.exe", "", "C:\Arquivos de programas\Microsoft Games\Age of Mythology", SW_SHOW)
        VGStrExecJogo = "aom.exe"

    End If

End Sub

Private Sub Timer1_Timer()
    
    Dim RecCred As ADODB.Recordset
    Dim RecMaq As ADODB.Recordset
    Dim RecConect As ADODB.Recordset
   
    hora_atual = Now
    
    LblDataHoraAtual.Caption = Now
    
    qtde = hora_atual - hora_inicial
    
    If qtde >= tempo_user Then  'tempo de crédito zerou
        
        'MsgBox ("Tempo do usuário esgotado")
        'Unload Me
      
        Conecta

        '======= Grava no banco as informações de tempo restante (zerado) ======================
        StrSql = "Update tb_credito set TempoRest='00:00:00' where NumCartao=" & VGIntNumCartao
        Set RecCred = vgCon.Execute(StrSql)
        '=========================================================================

        '======= Grava situação da máquina =============================================
        StrSql = "Update tb_maquina set Situacao='livre' where NumMaq=" & VGIntMaq
        Set RecMaq = vgCon.Execute(StrSql)
        '=========================================================================

        '======= Exclui usuário da lista de conexão =============================================
        StrSql = "Delete from tb_conect where CodCli=" & VGIntCodCli
        Set RecConect = vgCon.Execute(StrSql)
        '=========================================================================

        'MsgBox "Seu tempo acabou", vbOKOnly + vbCritical, "Atenção"

        Desconecta

        '============ Fecha o jogo ===================
        KillProgramInMemory VGStrExecJogo
        KillProgramInMemory VGStrExecJogo1
        '==========================================

        '=== Desabilita teclas de atalho ==================
        DisableCtrlAltDelete (True)
        '================================================

        '======= Reinicia computador forçado ===================
        RET = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
        '==========================================

        '======= Reinicia computador ===================
        RET = ExitWindowsEx(EWX_REBOOT, 0)
        '==========================================

        '======= Dar logoff do usuário ===================
        'RET = ExitWindowsEx(EWX_LOGOFF, 0)
        '==========================================

        '======= Desliga o computador ==================
        'RET = ExitWindowsEx(EWX_SHUTDOWN, 0)
        '==========================================

        Unload Me

        FrmIdentifica.Show
    Else
    
        restante = tempo_user - qtde
        LblResCredRest.Caption = restante
    
    End If
    '============================================================
    
    '======= Fechar aplicativos das máquinas pelo servidor ===================
    
    Dim VarMaq As Integer
    Dim RecCmd As ADODB.Recordset
    Dim RecDel As ADODB.Recordset
    'Dim RET As String

    VarMaq = Calcula_Maq(GetIPAddress())

    '========== ler tabela de comandos

    Conecta

    StrSql = "Select Maq from tb_comandos where Maq=" & VarMaq
    Set RecCmd = vgCon.Execute(StrSql)

    If Not RecCmd.EOF Then  'existe máquina na tabela
        StrSql = "Delete from tb_comandos where Maq=" & VarMaq
        Set RecDel = vgCon.Execute(StrSql)

        Desconecta

      '============ Fecha o jogo ===================
      KillProgramInMemory VGStrExecJogo
      KillProgramInMemory VGStrExecJogo1
      '==========================================

      Me.CmdSair.Value = True

      '======= Reinicia computador forçado ===================
      RET = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
      '==========================================

      '======= Reinicia computador ===================
      RET = ExitWindowsEx(EWX_REBOOT, 0)
      '==========================================

    End If


        'Desconecta

        'RET = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)

        'RET = ExitWindowsEx(EWX_REBOOT, 0)
    
    
    '============================================================
    
End Sub

Private Sub Timer2_Timer()
    CmdSair.Visible = True
End Sub

Private Sub Timer3_Timer()
    
    Dim RecCred As ADODB.Recordset
    
    Conecta
    
    StrSql = "Update tb_credito set TempoRest='" & LblResCredRest.Caption & "' where NumCartao=" & VGIntNumCartao
    Set RecCred = vgCon.Execute(StrSql)
        
    Desconecta
    'RecCred.Close
    
End Sub
