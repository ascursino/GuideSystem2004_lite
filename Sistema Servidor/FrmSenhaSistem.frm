VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmSenhaSistem 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Alteração das Senhas do Sistema"
   ClientHeight    =   3840
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9585
   Icon            =   "FrmSenhaSistem.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3840
   ScaleWidth      =   9585
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4560
      OleObjectBlob   =   "FrmSenhaSistem.frx":000C
      Top             =   2640
   End
   Begin VB.Frame FraAcesso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Alteração"
      ForeColor       =   &H80000008&
      Height          =   3615
      Left            =   6120
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtDescr 
         Appearance      =   0  'Flat
         Height          =   885
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Descrição da senha"
         Top             =   600
         Width           =   2895
      End
      Begin VB.TextBox TxtSenha 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   3
         ToolTipText     =   "Senha"
         Top             =   2520
         Width           =   2895
      End
      Begin VB.CommandButton CmdAlterar 
         Caption         =   "Alterar"
         Height          =   375
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Altera dados do acesso"
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox TxtLogin 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Login"
         Top             =   1920
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDescr 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSenhaSistem.frx":0240
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblLogin 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSenhaSistem.frx":02B0
         TabIndex        =   7
         Top             =   1680
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenha 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmSenhaSistem.frx":031A
         TabIndex        =   0
         Top             =   2280
         Width           =   615
      End
   End
   Begin FPSpread.vaSpread GrdSenha 
      Height          =   3615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5765
      _Version        =   393216
      _ExtentX        =   10169
      _ExtentY        =   6376
      _StockProps     =   64
      ColHeaderDisplay=   0
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GrayAreaBackColor=   16777215
      GridColor       =   0
      MaxCols         =   3
      MaxRows         =   0
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmSenhaSistem.frx":0384
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmSenhaSistem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public MsgLogin As String
Public MsgSenha As String
Public MsgDescr As String
Public VPStrBox As String
Public VPStrResponse As String
Public VPIntCodContr As Integer

Private Sub CmdAlterar_Click()
    Screen.MousePointer = vbHourglass
    
    If TxtDescr.Text = "" Or TxtLogin.Text = "" Or TxtSenha.Text = "" Then
        VPStrBox = MsgBox("Preencha os campos em branco", vbCritical, "Guide System - Aviso de erro")
    Else
        Conecta
        
        Dim RecPesq As New ADODB.Recordset
        Dim RecVerif As New ADODB.Recordset
        Dim RecLog As New ADODB.Recordset
                           
        StrSql = "Select Login,Senha from tb_controle where Login='" & TxtLogin.Text & "' or Senha='" & TxtSenha.Text & "'"
        RecPesq.Open StrSql, vgCon, 1, 3
               
        If Not RecPesq.EOF Then      'achou registro
            
            VPStrBox = MsgBox("Login ou Senha já existe.", vbCritical, "Guide System - Aviso de erro")
            TxtLogin.SetFocus
            Desconecta
        Else
            
            StrSql = "Select * from tb_controle where CodControle=" & VPIntCodContr
            RecVerif.Open StrSql, vgCon, 1, 3
            
            RecVerif("Descricao") = TxtDescr.Text
            RecVerif("Login") = TxtLogin.Text
            RecVerif("Senha") = TxtSenha.Text
            RecVerif.Update
            
            VPStrBox = MsgBox("Acesso modificado.", vbInformation, "Guide System - Informação")
              
            Desconecta
            Me.MontaGridSenha
        End If
                    
    End If
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmSenhaSistem.hwnd)
    
    Height = 4335
    Width = 9945
    'Top = 2070
    'Left = 1170
    
    Me.MontaGridSenha
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub TxtCodCli_KeyPress(KeyAscii As Integer)
    
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

Private Sub Form_Resize()
  FrmSenhaSistem.Left = (MDIPrincipal.Width / 2) - (FrmSenhaSistem.Width / 1.93)
  FrmSenhaSistem.Top = (MDIPrincipal.Height / 3) - (FrmSenhaSistem.Height / 5)
End Sub

Private Sub GrdSenha_DblClick(ByVal Col As Long, ByVal Row As Long)
    'Cód. Controle
    GrdSenha.Row = Row
    GrdSenha.Col = 4
    VPIntCodContr = GrdSenha.Text
    
    'Descrição
    GrdSenha.Row = Row
    GrdSenha.Col = 1
    TxtDescr.Text = GrdSenha.Text
    
    'Login
    GrdSenha.Row = Row
    GrdSenha.Col = 2
    TxtLogin.Text = GrdSenha.Text
    
    'Senha
    GrdSenha.Row = Row
    GrdSenha.Col = 3
    TxtSenha.Text = GrdSenha.Text
    
End Sub

Private Sub TxtDescr_GotFocus()
    TxtDescr.SelStart = 0
    TxtDescr.SelLength = Len(TxtDescr.Text)
End Sub

Private Sub TxtDescr_KeyPress(KeyAscii As Integer)
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

Private Sub TxtLogin_GotFocus()
    TxtLogin.SelStart = 0
    TxtLogin.SelLength = Len(TxtLogin.Text)
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

Private Sub TxtSenhaAtual_KeyPress(KeyAscii As Integer)
    
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

Private Sub TxtSenhaNova_KeyPress(KeyAscii As Integer)
    
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

Sub MontaGridSenha()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    Dim RecResult As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    
    StrSql = "Select CodControle,Descricao,Login,Senha from tb_controle"
    RecResult.Open StrSql, vgCon, 1, 3
        
    'If RecResult.EOF Then
    '       VPStrBox = MsgBox("Lista de espera vazia.", vbInformation, "Guide System - Informação")
    'End If
   
    VPIntLinha = 1
    
    GrdSenha.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        
        GrdSenha.Row = VPIntLinha
        GrdSenha.Lock = True
                        
        GrdSenha.Col = 1
        GrdSenha.Text = RecResult.Fields.Item(1).Value
        GrdSenha.Lock = True
        
        GrdSenha.Col = 2
        GrdSenha.Text = RecResult.Fields.Item(2).Value
        GrdSenha.Lock = True
        
        GrdSenha.Col = 3
        GrdSenha.Text = RecResult.Fields.Item(3).Value
        GrdSenha.Lock = True
        
        GrdSenha.Col = 4
        GrdSenha.Text = Val(RecResult.Fields.Item(0).Value)
        GrdSenha.Lock = True
        
        VPIntLinha = VPIntLinha + 1
        
        GrdSenha.MaxRows = GrdSenha.MaxRows + 1
        RecResult.MoveNext
    Loop
    GrdSenha.MaxRows = GrdSenha.MaxRows - 1
    RecResult.Close
    
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub TxtSenha_GotFocus()
    TxtSenha.SelStart = 0
    TxtSenha.SelLength = Len(TxtSenha.Text)
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
