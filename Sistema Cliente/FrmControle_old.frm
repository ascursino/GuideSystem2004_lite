VERSION 5.00
Begin VB.Form FrmControle 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tribus Victrix (Sistema Cliente) - Controle"
   ClientHeight    =   1965
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
   ScaleHeight     =   1965
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   30000
      Left            =   3480
      Top             =   840
   End
   Begin VB.Timer Timer1 
      Left            =   4080
      Top             =   840
   End
   Begin VB.CommandButton CmdSair 
      Caption         =   "Sair"
      Height          =   255
      Left            =   4080
      TabIndex        =   12
      ToolTipText     =   "Sai do sistema"
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label LblTempo 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "(00)"
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
      Left            =   1800
      TabIndex        =   13
      Top             =   1200
      Width           =   300
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
      TabIndex        =   11
      Top             =   960
      Width           =   795
   End
   Begin VB.Label LblResMaq 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00"
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   3120
      TabIndex        =   10
      Top             =   960
      Width           =   180
   End
   Begin VB.Label LblResHoraEntr 
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
      ForeColor       =   &H00808000&
      Height          =   210
      Left            =   1320
      TabIndex        =   9
      Top             =   1680
      Width           =   630
   End
   Begin VB.Label LblResDataEntr 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "00/00/0000"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   1440
      Width           =   810
   End
   Begin VB.Label LblResCredRest 
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
      ForeColor       =   &H00808000&
      Height          =   210
      Left            =   960
      TabIndex        =   7
      Top             =   1200
      Width           =   630
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
      TabIndex        =   6
      Top             =   960
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
      TabIndex        =   5
      Top             =   720
      Width           =   1020
   End
   Begin VB.Label LblHoraEntr 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Hora Entrada:"
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
      TabIndex        =   4
      Top             =   1680
      Width           =   1080
   End
   Begin VB.Label LblDataEntr 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Data Entrada:"
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
      Top             =   1440
      Width           =   1050
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
      Top             =   1200
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
      Top             =   960
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
      Top             =   720
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
Public HoraInicial As String
Public HoraFinal As String
Public VPIntHora As Integer
Public VPIntMin As Integer
Public VPIntSeg As Integer
Public VPIntHoraA As String
Public VPIntMinA As String
Public VPIntSegA As String
Public VPIntHoraRes As Integer
Public VPIntMinRes As Integer
Public VPIntSegRes As Integer
Public VPStrHoraSaida As String
Public VPStrBox As String
Public TempMin As Double
'Public TempMin2 As Double
Public TempHora As Double
Public RestTempo As Double
Public Contador As Double

Private Sub CmdSair_Click()
    
    'grava no banco informações de tempo restante
    
    Dim RecCred As ADODB.Recordset
    Dim RecMaq As ADODB.Recordset
    Dim RecMaqCli As ADODB.Recordset
    Dim RecConect As ADODB.Recordset
    Dim RET As Long
    
    ''RestTempo = Contador / 3600  '3600s = 1h    transforma segundos em horas 14,15
    
    ''If InStr(CStr(RestTempo), ",") Then
     ''   VPIntMin = Mid(RestTempo, (Len(InStr(CStr(RestTempo), ",")) + 1), 2)
      ''  If VPIntMin > 59 Then
            VPIntMin = 59
       '' Else
        ''    VPIntMin = VPIntMin
       '' End If
        
       '' VPIntHora = Mid(RestTempo, 1, (Len(InStr(RestTempo, ","))))
       '' If VPIntHora = "" Then
      ''      VPIntHora = 59
      ''  Else
      ''      VPIntHora = VPIntHora
      ''  End If
   '' Else
    ''    VPIntHora = RestTempo
   '' End If

    If Len(HoraFinal) = 19 Then
        VPIntHoraRes = Mid(HoraFinal, 12, 2)
        VPIntMinRes = Mid(HoraFinal, 15, 2)
        VPIntSegRes = Mid(HoraFinal, 18, 2)
    Else
        VPIntHoraRes = Mid(HoraFinal, 1, 2)
        VPIntMinRes = Mid(HoraFinal, 4, 2)
        VPIntSegRes = Mid(HoraFinal, 7, 2)
    
    End If
    
    VPStrHoraSaida = Time
    
    VGStrCredRest = TimeSerial(Hour(VPStrHoraSaida) - VPIntHoraRes, Minute(VPStrHoraSaida) - VPIntMinRes, Second(VPStrHoraSaida) - VPIntSegRes)
    'VGStrCredRest = VPIntHora & ":" & VPIntMin & ":00"
    
    Conecta
    
    StrSql = "Update tb_credito set TempoRest='" & VGStrCredRest & "' where NumCartao=" & VGIntNumCartao
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
    'VGStrCredRest = ""
    VGStrDataEntr = ""
    VGStrHoraEntr = ""
    VGStrDataSaida = ""
    VGStrHoraSaida = ""
    
    HoraInicial = ""
    HoraFinal = ""
    VPIntHora = 0
    VPIntMin = 0
    VPIntSeg = 0
    VPIntHoraRes = 0
    VPIntMinRes = 0
    VPIntSegRes = 0
    
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
    
    VPStrBox = MsgBox("Seu crédito restante é de " & VGStrCredRest, vbInformation, "Informação")
    VGStrCredRest = ""
    Unload FrmTela
    FrmIdentifica.Show
End Sub

Private Sub Form_Load()
    Height = 2340
    Width = 4890
    Top = 2880
    Left = 3555
    
    LblTempo.Visible = False
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
    '
    WinClass = "Shell_TrayWnd"
    '
    TaskBarHwnd = FindWindow(WinClass, vbNullString)
    '
    lRet = EnumChildWindows(TaskBarHwnd, AddressOf EnumChildProc, lParam)
    '=============================================
    
    '=========== Libera mouse no desktop ====================
    'Dim DesktopWindow As Retang
    
    'GetWindowRect GetDesktopWindow(), DesktopWindow
    'ClipCursor DesktopWindow
    '=============================================
    
    VGStrCredRest = VGStrCredito   'tempo de crédito
    Timer1.Enabled = True
    Timer1.Interval = 1000
    HoraInicial = Time
    
    VPIntHora = Mid(VGStrCredRest, 1, 2)
    VPIntMin = Mid(VGStrCredRest, 4, 2)
    VPIntSeg = Mid(VGStrCredRest, 7, 2)
    
    ''Contador = VPIntSeg
    
    ''TempMin = VPIntMin * 60     'transforma minutos em segundos
    
    ''Contador = Contador + TempMin
    
    ''TempHora = VPIntHora * 60    'transforma hora em minutos
    ''TempMin2 = TempHora * 60    'transforma minutos em segundos
    
    ''Contador = Contador + TempMin2     'guarda tempo do usuário em segundos
    
    VPIntHoraA = Mid(HoraInicial, 1, 2)
    VPIntMinA = Mid(HoraInicial, 4, 2)
    VPIntSegA = Mid(HoraInicial, 7, 2)
    
    HoraFinal = TimeSerial(Hour(HoraInicial) + VPIntHora, Minute(HoraInicial) + VPIntMin, Second(HoraInicial) + VPIntSeg)
    'HoraFinal = VPIntHoraA + VPIntHora & ":" & VPIntMinA + VPIntMin & ":" & VPIntSegA + VPIntSeg
        
    VGStrCredRest = VGStrCredito
    
    LblResCliente.Caption = FrmCliente.LblResNome
    LblResCartao.Caption = FormataNum(VGIntNumCartao)
    LblResMaq.Caption = FormataNum(VGIntMaq)
    LblResCredRest.Caption = VGStrCredRest
    LblResDataEntr = VGStrDataEntr
    LblResHoraEntr = HoraInicial
    LblTempo.Caption = "(" & Contador & ")"
    'LblResHoraSai = HoraFinal
    
    Dim i&
    
    If VGStrJogo = "strike" Then
        VGStrJogo = ""
        
        'inicia jogo Counter Strike
        'i& = ShellExecute(0, "open", "C:\windows\sol.exe", "", "C:\windows", SW_SHOW)
        i& = ShellExecute(0, "open", "C:\Sierra\Half-Life\hl.exe", "-console -game cstrike", "C:\Sierra\Half-Life", SW_SHOW)
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
        i& = ShellExecute(0, "open", "C:\AOE2CONQ\age2_x1.exe", "", "C:\AOE2CONQ", SW_SHOW)
        VGStrExecJogo = "age2_x1.exe"
        
    End If

End Sub

Private Sub Timer1_Timer()

    ''Contador = Contador - 1
    ''LblTempo.Caption = "(" & Contador & ")"
    
    '=============== Tempo de crédito esgotado ===================
    
    Dim RecCred As ADODB.Recordset
    Dim RecMaq As ADODB.Recordset
    Dim RecConect As ADODB.Recordset
    
    Hora = Mid(VGStrCredRest, 1, 2)
    Min = Mid(VGStrCredRest, 4, 2)
    Seg = Mid(VGStrCredRest, 1, 2)
    
    ''LblResCredRest.Caption = VGStrCredRest
    
    ''If Contador = 0 Then    'tempo de crédito zerou
    If Time >= HoraFinal Then  'tempo de crédito zerou
      LblResCredRest.Caption = VGStrCredRest
      
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
