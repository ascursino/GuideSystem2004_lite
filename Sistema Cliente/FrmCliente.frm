VERSION 5.00
Begin VB.Form FrmCliente 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tribus Victrix (Sistema Cliente) - Jogo"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.CommandButton CmdSair 
      Caption         =   "S A I R"
      Height          =   615
      Left            =   8400
      TabIndex        =   14
      TabStop         =   0   'False
      ToolTipText     =   "Sai do sistema"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton CmdIniciar 
      Caption         =   " I N I C I A R  J O G O "
      Height          =   615
      Left            =   2520
      TabIndex        =   10
      TabStop         =   0   'False
      ToolTipText     =   "Inicia jogo selecionado"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.Frame FraControle 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Controle"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2160
      TabIndex        =   2
      Top             =   2880
      Width           =   7695
      Begin VB.Label LblResTempo 
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
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   630
      End
      Begin VB.Label LblTempo 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tempo:"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   540
      End
      Begin VB.Label LblCred 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Crédito"
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
         Left            =   120
         TabIndex        =   8
         Top             =   720
         Width           =   645
      End
      Begin VB.Label LblAnos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "anos"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   7080
         TabIndex        =   7
         Top             =   360
         Width           =   345
      End
      Begin VB.Label LblResIdade 
         BackColor       =   &H00FFFFFF&
         Caption         =   "00"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   6840
         TabIndex        =   6
         Top             =   360
         Width           =   285
      End
      Begin VB.Label LblResNome 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "(nome do cliente)"
         ForeColor       =   &H00808000&
         Height          =   195
         Left            =   840
         TabIndex        =   5
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label LblIdade 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Idade:"
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
         Left            =   6180
         TabIndex        =   4
         Top             =   360
         Width           =   585
      End
      Begin VB.Label LblNome 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cliente:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   675
      End
   End
   Begin VB.Frame FraJogos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Jogos disponíveis"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   4800
      Width           =   11535
      Begin VB.Image ImgAgeMit 
         Height          =   705
         Left            =   10440
         Picture         =   "FrmCliente.frx":000C
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblAgeMit 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Age of Mythology"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   10200
         TabIndex        =   19
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Shape ShpAgeMit 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   10200
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape ShpAge 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   8760
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblAge 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Age of Empires"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   8835
         TabIndex        =   18
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image ImgAge 
         Height          =   705
         Left            =   9000
         Picture         =   "FrmCliente.frx":0460
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Image ImgBat 
         Height          =   705
         Left            =   7560
         Picture         =   "FrmCliente.frx":082B
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblBat 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "BattleField"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   7560
         TabIndex        =   17
         Top             =   1080
         Width           =   765
      End
      Begin VB.Shape ShpBat 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   7320
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image ImgDay 
         Height          =   705
         Left            =   6120
         Picture         =   "FrmCliente.frx":1D0D
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblDay 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Day of Defeat"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6000
         TabIndex        =   16
         Top             =   1080
         Width           =   1005
      End
      Begin VB.Shape ShpDay 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   5880
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Image ImgDeath 
         Height          =   705
         Left            =   4680
         Picture         =   "FrmCliente.frx":31EF
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblDeath 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "DeathMatch"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4605
         TabIndex        =   15
         Top             =   1080
         Width           =   915
      End
      Begin VB.Shape ShpDeath 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   4440
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape ShpWar 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   3000
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape ShpHalf 
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   1560
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Shape ShpStrike 
         BackColor       =   &H00FFFFFF&
         BorderColor     =   &H000000C0&
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label LblWar 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "WarCraft III"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3210
         TabIndex        =   12
         Top             =   1080
         Width           =   825
      End
      Begin VB.Image ImgWar 
         Height          =   705
         Left            =   3240
         Picture         =   "FrmCliente.frx":46D1
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblHalf 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "HalfLife"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1860
         TabIndex        =   11
         Top             =   1080
         Width           =   555
      End
      Begin VB.Image ImgHalf 
         Height          =   705
         Left            =   1800
         Picture         =   "FrmCliente.frx":5BB3
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LblStrike 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Counter Strike"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   1155
      End
      Begin VB.Image ImgStrike 
         Height          =   705
         Left            =   360
         Picture         =   "FrmCliente.frx":7095
         Stretch         =   -1  'True
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Image ImgTribus 
      Height          =   9090
      Left            =   0
      Picture         =   "FrmCliente.frx":8577
      Stretch         =   -1  'True
      Top             =   0
      Width           =   12090
   End
End
Attribute VB_Name = "FrmCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdIniciar_Click()

    VGStrDataEntr = FormataData(Date)
    VGStrHoraEntr = FormataHora(Time)
    'VGIntMaq = Calcula_Maq(GetIPAddress())
    
    If VGStrIP = "inexistente" Then
        VGStrIP = ""
        VGIntNumCartao = 0
        VGStrCredito = ""
        VGIntCodCli = 0
        VGStrDataEntr = ""
        VGStrDataSaida = ""
        VGStrHoraEntr = ""
        VGStrHoraSaida = ""
        VGIntMaq = 0
        VGIntIdade = 0
        
        Unload Me
        FrmIdentifica.Show
        
    Else
    
        Conecta
        
        Dim RecMaq As ADODB.Recordset
        Dim RecMaqCli As ADODB.Recordset
        Dim RecConect As ADODB.Recordset
                           
        StrSql = "Update tb_maquina set Situacao='ocupado' where NumMaq=" & VGIntMaq
        Set RecMaq = vgCon.Execute(StrSql)
                           
        'StrSql = "Insert into tb_maquina values (" & VGIntMaq & ",'ocupado')"
        'Set RecMaq = vgCon.Execute(StrSql)
        
        StrSql = "Insert into tb_maqcli values (" & VGIntCodCli & "," & VGIntMaq & ",'" & FormataDataUS(VGStrDataEntr) & "','" & VGStrHoraEntr & "',null,null)"
        Set RecMaqCli = vgCon.Execute(StrSql)
        
        Desconecta
        
        Unload Me
        
        FrmTela.Show
        FrmControle.Show
    
    End If

End Sub

Private Sub CmdSair_Click()
    Conecta
     
    Dim RecConect As ADODB.Recordset

    StrSql = "Delete from tb_conect where CodCli=" & VGIntCodCli
    Set RecConect = vgCon.Execute(StrSql)
    
    Desconecta
    
    Unload Me
    FrmIdentifica.Show
     
End Sub

Private Sub Form_Load()
    Height = 9090
    Width = 12090
    'Top = 1590
    'Left = 1890
    
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    CmdIniciar.Enabled = False
    
    '=== Desabilita teclas de atalho ==================
    DisableCtrlAltDelete (True)
    '================================================

    '====== Não consegue desligar o sistema via ctrl+alt+del =====
    Call MakeMeService
    '===============================================

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

    '=============== Prende mouse no form ===============
    'Dim EstaJanela As Retang

    'GetWindowRect Me.hWnd, EstaJanela
    'ClipCursor EstaJanela
    '===============================================

    Conecta

    Dim RecCli As ADODB.Recordset

    StrSql = "Select Nome,NascDia,NascMes,NascAno from tb_cliente where CodCli=" & VGIntCodCli
    Set RecCli = vgCon.Execute(StrSql)

    LblResNome.Caption = RecCli.Fields.Item(0).Value
    LblResIdade.Caption = Calcula_Idade(RecCli.Fields.Item(1).Value, RecCli.Fields.Item(2).Value, RecCli.Fields.Item(3).Value)
    LblResTempo.Caption = VGStrCredito

    VGIntIdade = LblResIdade

    Desconecta

    If VGStrRestIdade = "sim" Then
        Call AcessoJogo
        'VGStrRestIdade = ""
    End If
    
End Sub

Private Sub ImgBat_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = True
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "bat"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub ImgDeath_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = True
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "death"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub ImgHalf_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = True
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "half"
    
    CmdIniciar.Enabled = True
    
End Sub

Private Sub ImgDay_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = True
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "day"
    
    CmdIniciar.Enabled = True
    
End Sub

Private Sub ImgStrike_Click()
    ShpStrike.Visible = True
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "strike"
    
    CmdIniciar.Enabled = True
End Sub

Private Sub ImgWar_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = True
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "war"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub ImgAge_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = True
    ShpAgeMit.Visible = False
    
    VGStrJogo = "age"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub ImgAgeMit_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = True
    
    VGStrJogo = "agemit"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub LblBat_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = True
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "bat"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub LblDeath_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = True
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "death"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub LblHalf_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = True
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "half"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub LblDay_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = True
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "day"
    
    CmdIniciar.Enabled = True
    
End Sub

Private Sub LblStrike_Click()
    ShpStrike.Visible = True
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "strike"
    
    CmdIniciar.Enabled = True
    
End Sub

Private Sub LblAge_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = True
    ShpAgeMit.Visible = False
    
    VGStrJogo = "age"
    
    CmdIniciar.Enabled = True
    
End Sub

Private Sub LblWar_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = True
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = False
    
    VGStrJogo = "war"
    
    CmdIniciar.Enabled = True

End Sub

Private Sub LblAgeMit_Click()
    ShpStrike.Visible = False
    ShpHalf.Visible = False
    ShpWar.Visible = False
    ShpDeath.Visible = False
    ShpDay.Visible = False
    ShpBat.Visible = False
    ShpAge.Visible = False
    ShpAgeMit.Visible = True
    
    VGStrJogo = "agemit"
    
    CmdIniciar.Enabled = True

End Sub

Sub AcessoJogo()
'========= Verifica faixa etária para os jogos =========

'============== Counter Strike ==================
    If VGIntIdade < 18 Then     'não habilita jogo
        ImgStrike.Enabled = False
        LblStrike.Enabled = False
        LblStrike.ForeColor = &H808080
        
    ElseIf VGIntIdade >= 18 Then    'habilita jogo
        ImgStrike.Enabled = True
        LblStrike.Enabled = True
    
    End If
'===========================================

'================== Half Life ==================
    If VGIntIdade < 18 Then     'não habilita jogo
        ImgHalf.Enabled = False
        LblHalf.Enabled = False
        LblHalf.ForeColor = &H808080
        
    ElseIf VGIntIdade >= 18 Then    'habilita jogo
        ImgHalf.Enabled = True
        LblHalf.Enabled = True
    
    End If
'===========================================

'================= WarCraft III =================
    If VGIntIdade < 14 Then     'não habilita jogo
        ImgWar.Enabled = False
        LblWar.Enabled = False
        LblWar.ForeColor = &H808080
        
    ElseIf VGIntIdade >= 14 Then    'habilita jogo
        ImgWar.Enabled = True
        LblWar.Enabled = True
    
    End If
'==========================================

'================ DeathMatch =================
    If VGIntIdade < 18 Then     'não habilita jogo
        ImgDeath.Enabled = False
        LblDeath.Enabled = False
        LblDeath.ForeColor = &H808080
        
    ElseIf VGIntIdade >= 18 Then    'habilita jogo
        ImgDeath.Enabled = True
        LblDeath.Enabled = True
    
    End If
'==========================================

'============== Day of Defeat (DOD) =============
    If VGIntIdade < 16 Then     'não habilita jogo
        ImgDay.Enabled = False
        LblDay.Enabled = False
        LblDay.ForeColor = &H808080
        
    ElseIf VGIntIdade >= 16 Then    'habilita jogo
        ImgDay.Enabled = True
        LblDay.Enabled = True
    
    End If
'==========================================
 
'============== Battle Field 1942 ===============
    If VGIntIdade < 16 Then     'não habilita jogo
        ImgBat.Enabled = False
        LblBat.Enabled = False
        LblBat.ForeColor = &H808080
        
    ElseIf VGIntIdade >= 16 Then    'habilita jogo
        ImgBat.Enabled = True
        LblBat.Enabled = True
    
    End If
'==========================================
 
'============== Age Of Empires ===============
'    If VGIntIdade < 16 Then     'não habilita jogo
'        ImgBat.Enabled = False
'        LblBat.Enabled = False
'        LblBat.ForeColor = &H808080
'
'    ElseIf VGIntIdade >= 16 Then    'habilita jogo
'        ImgBat.Enabled = True
'        LblBat.Enabled = True
'
'    End If
'==========================================

'============== Age Of Mythology ===============
'    If VGIntIdade < 16 Then     'não habilita jogo
'        ImgBat.Enabled = False
'        LblBat.Enabled = False
'        LblBat.ForeColor = &H808080
'
'    ElseIf VGIntIdade >= 16 Then    'habilita jogo
'        ImgBat.Enabled = True
'        LblBat.Enabled = True
'
'    End If
'==========================================

End Sub

