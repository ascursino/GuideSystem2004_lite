VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCadCli 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "FrmCadCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   6630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraCadCli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4770
      Left            =   120
      TabIndex        =   16
      Top             =   0
      Width           =   6135
      Begin VB.TextBox TxtEstado 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   2
         TabIndex        =   6
         ToolTipText     =   "Estado"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox TxtDtCad 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Text            =   "__/__/____"
         ToolTipText     =   "Data do Cadastro"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox TxtCep 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   3
         ToolTipText     =   "Cep"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtCidade 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   100
         TabIndex        =   5
         ToolTipText     =   "Cidade"
         Top             =   1800
         Width           =   1695
      End
      Begin VB.TextBox TxtBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   4
         ToolTipText     =   "Bairro"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox TxtCpf 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   12
         TabIndex        =   13
         ToolTipText     =   "CPF"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox TxtTelRec 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   9
         ToolTipText     =   "Telefone de recado"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtCel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   9
         TabIndex        =   8
         ToolTipText     =   "Celular"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton CmdIncluir 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   4800
         TabIndex        =   15
         ToolTipText     =   "Inclui cadastro"
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox TxtObs 
         Appearance      =   0  'Flat
         Height          =   675
         Left            =   240
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   14
         ToolTipText     =   "Observação sobre o cliente"
         Top             =   3960
         Width           =   4335
      End
      Begin VB.TextBox TxtIdent 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   20
         TabIndex        =   12
         ToolTipText     =   "Identidade"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox TxtDtNasc 
         Appearance      =   0  'Flat
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   3
         EndProperty
         Height          =   285
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "__/__/____"
         ToolTipText     =   "Data de nascimento"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox TxtContato 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4200
         MaxLength       =   50
         TabIndex        =   10
         ToolTipText     =   "Pessoa de contato"
         Top             =   2520
         Width           =   1695
      End
      Begin VB.TextBox TxtTelRes 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   9
         TabIndex        =   7
         ToolTipText     =   "Telefone de residência"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox TxtEnd 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   250
         TabIndex        =   2
         ToolTipText     =   "Endereço"
         Top             =   1080
         Width           =   4575
      End
      Begin VB.TextBox TxtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         MaxLength       =   150
         TabIndex        =   1
         ToolTipText     =   "Nome completo"
         Top             =   720
         Width           =   4575
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   5040
         OleObjectBlob   =   "FrmCadCli.frx":000C
         Top             =   3600
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCad 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":0240
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNome 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":02C0
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEnd 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":0328
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCep 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":0398
         TabIndex        =   20
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBairro 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCadCli.frx":03FE
         TabIndex        =   21
         Top             =   1440
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCidade 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":046A
         TabIndex        =   22
         Top             =   1800
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblEstado 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCadCli.frx":04D6
         TabIndex        =   23
         Top             =   1800
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTelRes 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":0542
         TabIndex        =   24
         Top             =   2160
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCel 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCadCli.frx":05B2
         TabIndex        =   25
         Top             =   2160
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblTelRec 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":0620
         TabIndex        =   26
         Top             =   2520
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblContato 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCadCli.frx":0696
         TabIndex        =   27
         Top             =   2520
         Width           =   735
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtNasc 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":0704
         TabIndex        =   28
         Top             =   2880
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblExDtNasc 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCadCli.frx":0776
         TabIndex        =   29
         Top             =   2880
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblIdent 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":07F2
         TabIndex        =   30
         Top             =   3240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCpf 
         Height          =   255
         Left            =   3360
         OleObjectBlob   =   "FrmCadCli.frx":0866
         TabIndex        =   31
         Top             =   3240
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblObs 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmCadCli.frx":08CC
         TabIndex        =   32
         Top             =   3600
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmCadCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgNome As String
Public MsgEnd As String
Public MsgCep As String
Public MsgBairro As String
Public MsgCidade As String
Public MsgEstado As String
Public MsgTelRes As String
Public MsgDtNasc As String
Public MsgIdent As String
Public MsgCpf As String
Public MsgPai As String
Public MsgMae  As String
Public MsgCpfPai As String
Public MsgCpfMae As String
Public VPStrBox As String
Public VPStrGrava As String
Public VPIntDia As Integer
Public VPIntMes As Integer
Public VPIntAno As Integer

Private Sub CmdIncluir_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtNome.Text = "" Then
        VPStrBox = MsgBox("Preencha pelo menos o campo nome", vbCritical, "Guide System - Aviso de erro")
    Else
        Conecta
        
        Dim RecCli As New ADODB.Recordset
        
        If TxtDtNasc.Text = "" Or TxtDtNasc.Text = "__/__/____" Then
            VPIntDia = 0
            VPIntMes = 0
            VPIntAno = 0
        Else
            VPIntDia = FormataNum(Mid(TxtDtNasc.Text, 1, 2))
            VPIntMes = FormataNum(Mid(TxtDtNasc.Text, 4, 2))
            VPIntAno = Mid(TxtDtNasc.Text, 7, 4)
        End If
        
        StrSql = "Select * from tb_cliente"
        RecCli.Open StrSql, vgCon, 1, 3
    
        RecCli.AddNew
        RecCli("DtCad") = FormataDataUS(TxtDtCad.Text)
        RecCli("Nome") = TxtNome.Text
        RecCli("Ender") = TxtEnd.Text
        RecCli("Cep") = TxtCep.Text
        RecCli("Bairro") = TxtBairro.Text
        RecCli("Cidade") = TxtCidade.Text
        RecCli("Estado") = TxtEstado.Text
        RecCli("Tel") = TxtTelRes.Text
        RecCli("Cel") = TxtCel.Text
        RecCli("TelRec") = TxtTelRec.Text
        RecCli("Contato") = TxtContato.Text
        RecCli("NascDia") = VPIntDia
        RecCli("NascMes") = VPIntMes
        RecCli("NascAno") = VPIntAno
        RecCli("Ident") = TxtIdent.Text
        RecCli("Cpf") = TxtCpf.Text
        RecCli("Obs") = TxtObs.Text
        RecCli.Update
          
        Desconecta
          
        VPStrBox = MsgBox("Cadastro efetuado.", vbInformation, "Guide System - Informação")
        
        VGStrResponse = MsgBox("Deseja cadastrar um" & Chr(13) & "acesso para este cliente?", vbYesNo)
        
        If VGStrResponse = vbYes Then
            VGStrForm = "Cliente"
            Call PegaCodCli
            FrmAcesso.Show
        Else
            Unload Me
        End If
        
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCadCli.hwnd)
    
    Height = 5400
    Width = 6720
    'Top = 1275
    'Left = 2475
    
    'Call PegaCodCli
    
    TxtDtCad.Text = FormataData(Date)
    TxtCidade.Text = "Rio de Janeiro"
    TxtEstado.Text = "RJ"
    
    Screen.MousePointer = vbNormal
End Sub

Sub PegaCodCli()

    Conecta

    Dim RecCod As New ADODB.Recordset
    StrSql = "Select max(CodCli) as CodCli From tb_cliente"
    RecCod.Open StrSql, vgCon, 1, 3

    If Not RecCod.EOF Then 'tem cadastrado
        VGIntCodCli = FormataNum(RecCod.Fields.Item(0).Value)
    End If

    RecCod.Close
    Desconecta

End Sub

Sub VerifBranco()

'    Dim Dia As Integer
'    Dim Mes As Integer
'    Dim Ano As Integer
'    Dim Idade As Integer
    
'    If Len(TxtDtNasc.Text) <> 10 Or TxtDtNasc.Text = "" Then
'        VPStrBox = MsgBox("Campo 'Data Nasc' não está no padrão exigido." & Chr(13) & "Ex: dd/mm/aaaa", vbCritical, "Guide System - Aviso de erro")
'    End If
    
'    If Len(TxtDtNasc.Text) = 10 And TxtDtNasc.Text <> "" Then
'        Dia = FormataNum(Mid(TxtDtNasc.Text, 1, 2))
'        Mes = FormataNum(Mid(TxtDtNasc.Text, 4, 2))
'        Ano = Mid(TxtDtNasc.Text, 7, 4)
'        Idade = Calcula_Idade(Dia, Mes, Ano)
'    Else
'        Idade = 0
'    End If
    
'    If TxtNome.Text = "" Or TxtEnd.Text = "" Or TxtCep.Text = "" Or TxtBairro.Text = "" Or TxtCidade.Text = "" Or TxtEstado.Text = "" Or TxtTelRes.Text = "" Or TxtDtNasc.Text = "" Or TxtIdent.Text = "" Or TxtCpf.Text = "" Or TxtPai.Text = "" Or TxtMae.Text = "" Or TxtCpfPai.Text = "" Or TxtCpfMae.Text = "" Then
        
'        If TxtNome.Text = "" Then
'            MsgNome = " - Nome" & Chr(13)
'        End If
        
'        If TxtEnd.Text = "" Then
'            MsgEnd = " - Endereço" & Chr(13)
'        End If
        
'        If TxtCep.Text = "" Then
'            MsgCep = " - Cep" & Chr(13)
'        End If
        
'        If TxtBairro.Text = "" Then
'            MsgBairro = " - Bairro" & Chr(13)
'        End If
        
'        If TxtCidade.Text = "" Then
'            MsgCidade = " - Cidade" & Chr(13)
'        End If
        
'        If TxtEstado.Text = "" Then
'            MsgEstado = " - Estado" & Chr(13)
'        End If
        
'        If TxtTelRes.Text = "" Then
'            MsgTelRes = " - Telefone" & Chr(13)
'        End If
        
'        If TxtDtNasc.Text = "" Then
'            MsgDtNasc = " - Data Nasc" & Chr(13)
'        End If
        
'        If Idade = 0 Then
            
'            If TxtIdent.Text = "" Then
'                MsgIdent = " - Identidade" & Chr(13)
'            End If
            
'            If TxtCpf.Text = "" Then
'                MsgCpf = " - Cpf" & Chr(13)
'            End If
        
'            If TxtPai.Text = "" Then
'                MsgPai = " - Pai" & Chr(13)
'            End If
            
'            If TxtMae.Text = "" Then
'                MsgMae = " - Mãe" & Chr(13)
'            End If
       
'            If TxtCpfPai.Text = "" Then
'                MsgCpfPai = " - Cpf do pai" & Chr(13)
'            End If
        
'            If TxtCpfMae.Text = "" Then
'                MsgCpfMae = " - Cpf da mãe" & Chr(13)
'            End If
        
'        ElseIf Idade >= 18 Then
        
'            If TxtIdent.Text = "" Then
'                MsgIdent = " - Identidade" & Chr(13)
'            End If
            
'            If TxtCpf.Text = "" Then
'                MsgCpf = " - Cpf" & Chr(13)
'            End If
        
'        ElseIf Idade <= 18 Then
            
'            If TxtPai.Text = "" Then
'                MsgPai = " - Pai" & Chr(13)
'            End If
            
'            If TxtMae.Text = "" Then
'                MsgMae = " - Mãe" & Chr(13)
'            End If
       
'            If TxtCpfPai.Text = "" Then
'                MsgCpfPai = " - Cpf do pai" & Chr(13)
'            End If
        
'            If TxtCpfMae.Text = "" Then
'                MsgCpfMae = " - Cpf da mãe" & Chr(13)
'            End If
        
'        End If
        
'        VPStrBox = MsgBox("Campo(s) não pode(m) estar em branco:" & Chr(13) & Chr(13) & MsgNome & MsgEnd & MsgCep & MsgBairro & MsgCidade & MsgEstado & MsgTelRes & MsgDtNasc & MsgIdent & MsgCpf & MsgPai & MsgMae & MsgCpfPai & MsgCpfMae, vbCritical, "Guide System - Aviso de erro")
        
'        MsgNome = ""
'        MsgEnd = ""
'        MsgCep = ""
'        MsgBairro = ""
'        MsgCidade = ""
'        MsgEstado = ""
'        MsgTelRes = ""
'        MsgDtNasc = ""
'        MsgIdent = ""
'        MsgCpf = ""
'        MsgPai = ""
'        MsgMae = ""
'        MsgCpfPai = ""
'        MsgCpfMae = ""
'    Else
        If Len(TxtDtNasc.Text) <> 10 Then
            VPStrBox = MsgBox("O campo 'Data Nasc' não possui" & Chr(13) & "o formato correto: dd/mm/yyyy", vbCritical, "Guide System - Aviso de erro")
        Else
             VPStrGrava = "Sim"
        End If
'    End If

End Sub

Private Sub Form_Resize()
  FrmCadCli.Left = (MDIPrincipal.Width / 2) - (FrmCadCli.Width / 2)
  FrmCadCli.Top = (MDIPrincipal.Height / 3) - (FrmCadCli.Height / 5)
End Sub

Private Sub TxtBairro_GotFocus()
    TxtBairro.SelStart = 0
    TxtBairro.SelLength = Len(TxtBairro.Text)
End Sub

Private Sub TxtBairro_KeyPress(KeyAscii As Integer)
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

Private Sub TxtCel_GotFocus()
    TxtCel.SelStart = 0
    TxtCel.SelLength = Len(TxtCel.Text)
End Sub

Private Sub TxtCel_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCep_GotFocus()
    TxtCep.SelStart = 0
    TxtCep.SelLength = Len(TxtCep.Text)
End Sub

Private Sub TxtCep_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtCodCli_GotFocus()
    TxtCodCli.SelStart = 0
    TxtCodCli.SelLength = Len(TxtCodCli.Text)
End Sub

Private Sub TxtCodCli_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtContato_GotFocus()
    TxtContato.SelStart = 0
    TxtContato.SelLength = Len(TxtContato.Text)
End Sub

Private Sub TxtContato_KeyPress(KeyAscii As Integer)
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

Private Sub TxtCpf_GotFocus()
    TxtCpf.SelStart = 0
    TxtCpf.SelLength = Len(TxtCpf.Text)
End Sub

Private Sub TxtCpf_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtCad_GotFocus()
    If TxtDtCad.Text = "__/__/____" Then
        TxtDtCad.Text = ""
    End If
    
    TxtDtCad.SelStart = 0
    TxtDtCad.SelLength = Len(TxtDtCad.Text)
End Sub

Private Sub TxtDtCad_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If TxtDtCad.Text = "__/__/____" Then
        TxtDtCad.Text = ""
    End If
End Sub

Private Sub TxtDtCad_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtCad.Text <> "" Then
        VLStrData = VerificaData(TxtDtCad.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtCad.SetFocus
        Else
            TxtDtCad.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtCad.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtNasc_GotFocus()
    If TxtDtNasc.Text = "__/__/____" Then
        TxtDtNasc.Text = ""
    End If
    
    TxtDtNasc.SelStart = 0
    TxtDtNasc.SelLength = Len(TxtDtNasc.Text)
End Sub

Private Sub TxtDtNasc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
    
    If TxtDtNasc.Text = "__/__/____" Then
        TxtDtNasc.Text = ""
    End If
End Sub

Private Sub TxtDtNasc_LostFocus()
    
    Dim VLStrData As String
    
    If TxtDtNasc.Text <> "" Then
        VLStrData = VerificaData(TxtDtNasc.Text)
        
        If VGStrDataErro = "sim" Then
            TxtDtNasc.SetFocus
        Else
            TxtDtNasc.Text = VLStrData
        End If
        
        VGStrDataErro = ""
    Else
        TxtDtNasc.Text = "__/__/____"
    End If
End Sub

Private Sub TxtEnd_GotFocus()
    TxtEnd.SelStart = 0
    TxtEnd.SelLength = Len(TxtEnd.Text)
End Sub

Private Sub TxtEnd_KeyPress(KeyAscii As Integer)
    '============ Símbolos ======================
    If (KeyAscii >= 33 And KeyAscii <= 43) Or KeyAscii = 46 Or KeyAscii = 47 Then
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

Private Sub TxtEstado_GotFocus()
    TxtEstado.SelStart = 0
    TxtEstado.SelLength = Len(TxtEstado.Text)
End Sub

Private Sub TxtEstado_KeyPress(KeyAscii As Integer)
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

Private Sub TxtIdent_GotFocus()
    TxtIdent.SelStart = 0
    TxtIdent.SelLength = Len(TxtIdent.Text)
End Sub

Private Sub TxtIdent_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtNome_GotFocus()
    TxtNome.SelStart = 0
    TxtNome.SelLength = Len(TxtNome.Text)
End Sub

Private Sub TxtNome_KeyPress(KeyAscii As Integer)
    
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

Private Sub TxtObs_GotFocus()
    TxtObs.SelStart = 0
    TxtObs.SelLength = Len(TxtObs.Text)
End Sub

Private Sub TxtObs_KeyPress(KeyAscii As Integer)
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

Private Sub TxtTelRec_GotFocus()
    TxtTelRec.SelStart = 0
    TxtTelRec.SelLength = Len(TxtTelRec.Text)
End Sub

Private Sub TxtTelRec_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtTelRes_GotFocus()
    TxtTelRes.SelStart = 0
    TxtTelRes.SelLength = Len(TxtTelRes.Text)
End Sub

Private Sub TxtTelRes_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

