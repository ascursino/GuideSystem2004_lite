VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmConsultCli 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Clientes"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   Icon            =   "FrmConsultCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   10200
      TabIndex        =   9
      ToolTipText     =   "Consulta movimento do caixa"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame FraConsult 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Consulta"
      ForeColor       =   &H80000008&
      Height          =   1815
      Left            =   840
      TabIndex        =   7
      Top             =   120
      Width           =   8415
      Begin VB.ComboBox CboDtCad 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCli.frx":000C
         Left            =   1320
         List            =   "FrmConsultCli.frx":000E
         TabIndex        =   8
         ToolTipText     =   "Bairro"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtIdent 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Identidade"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtCpf 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         MaxLength       =   12
         TabIndex        =   4
         ToolTipText     =   "CPF"
         Top             =   720
         Width           =   1575
      End
      Begin VB.TextBox TxtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   1
         ToolTipText     =   "Nome"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox TxtCodCli 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   0
         ToolTipText     =   "Código do cliente"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox CboBairro 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCli.frx":0010
         Left            =   3960
         List            =   "FrmConsultCli.frx":0012
         TabIndex        =   2
         ToolTipText     =   "Bairro"
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton CmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         ToolTipText     =   "Consulta clientes"
         Top             =   1200
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConsultCli.frx":0014
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCad 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConsultCli.frx":008A
         TabIndex        =   11
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNome 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmConsultCli.frx":0102
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblBairro 
         Height          =   255
         Left            =   3240
         OleObjectBlob   =   "FrmConsultCli.frx":0168
         TabIndex        =   13
         Top             =   720
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblIdent 
         Height          =   255
         Left            =   6000
         OleObjectBlob   =   "FrmConsultCli.frx":01D2
         TabIndex        =   14
         Top             =   360
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCpf 
         Height          =   255
         Left            =   6000
         OleObjectBlob   =   "FrmConsultCli.frx":023C
         TabIndex        =   15
         Top             =   720
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "FrmConsultCli.frx":02A0
      Top             =   720
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTotalCad 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "FrmConsultCli.frx":04D4
      TabIndex        =   6
      Top             =   2040
      Width           =   1575
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblQtde 
      Height          =   255
      Left            =   1920
      OleObjectBlob   =   "FrmConsultCli.frx":0558
      TabIndex        =   16
      Top             =   2040
      Width           =   975
   End
   Begin FPSpread.vaSpread GrdCli 
      Height          =   3135
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   9735
      _Version        =   393216
      _ExtentX        =   17171
      _ExtentY        =   5530
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
      MaxCols         =   17
      MaxRows         =   0
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmConsultCli.frx":05BE
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmConsultCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecResult As New ADODB.Recordset
Public VPStrBox As String
Public VPIntLinha As Integer

Private Sub CmdConsultar_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select * from tb_cliente where 0=0"
        
     If TxtCodCli.Text <> "" Then
        StrSql = StrSql + " and CodCli=" & TxtCodCli.Text & ""
     End If
            
     If CboDtCad.Text <> "" Then
        StrSql = StrSql + " and DtCad=#" & FormataDataUS(CboDtCad.Text) & "#"
     End If
       
     If TxtNome.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNome.Text & "%'"
     End If
            
     If CboBairro.Text <> "" Then
        StrSql = StrSql + " and Bairro='" & CboBairro.Text & "'"
     End If
            
     If TxtIdent.Text <> "" Then
        StrSql = StrSql + " and Ident='" & TxtIdent.Text & "'"
     End If
            
     If TxtCpf.Text <> "" Then
        StrSql = StrSql + " and Cpf='" & TxtCpf.Text & "'"
     End If
        
     StrSql = StrSql + " order by Nome asc"
     RecResult.Open StrSql, vgCon, 1, 3
        
     Call MontaGridCli
     
     Desconecta
     
     Screen.MousePointer = vbNormal
End Sub

Sub MontaGridCli()
    
    If RecResult.EOF Then
           VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Guide System - Informação")
    End If
   
    VPIntLinha = 1
    
    GrdCli.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        
        GrdCli.Row = VPIntLinha
        GrdCli.Lock = True
                        
        GrdCli.Col = 1
        GrdCli.Text = FormataNum(RecResult.Fields.Item(0).Value)
        GrdCli.Lock = True
        
        GrdCli.Col = 2
        GrdCli.Text = FormataData(RecResult.Fields.Item(1).Value)
        GrdCli.Lock = True
        
        GrdCli.Col = 3
        GrdCli.Text = RecResult.Fields.Item(2).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 4
        GrdCli.Text = RecResult.Fields.Item(3).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 5
        GrdCli.Text = RecResult.Fields.Item(4).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 6
        GrdCli.Text = RecResult.Fields.Item(5).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 7
        GrdCli.Text = RecResult.Fields.Item(6).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 8
        GrdCli.Text = RecResult.Fields.Item(7).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 9
        GrdCli.Text = RecResult.Fields.Item(8).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 10
        GrdCli.Text = RecResult.Fields.Item(9).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 11
        GrdCli.Text = RecResult.Fields.Item(10).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 12
        GrdCli.Text = RecResult.Fields.Item(11).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 13
        GrdCli.Text = FormataNum(RecResult.Fields.Item(12).Value) & "/" & FormataNum(RecResult.Fields.Item(13).Value) & "/" & RecResult.Fields.Item(14).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 14
        GrdCli.Text = Val(Calcula_Idade(RecResult.Fields.Item(12).Value, RecResult.Fields.Item(13).Value, RecResult.Fields.Item(14).Value))
        GrdCli.Lock = True
           
        GrdCli.Col = 15
        GrdCli.Text = RecResult.Fields.Item(15).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 16
        GrdCli.Text = RecResult.Fields.Item(16).Value
        GrdCli.Lock = True
           
        GrdCli.Col = 17
        GrdCli.Text = RecResult.Fields.Item(17).Value
        GrdCli.Lock = True
        
        VPIntLinha = VPIntLinha + 1
        
        GrdCli.MaxRows = GrdCli.MaxRows + 1
        RecResult.MoveNext
    Loop
    GrdCli.MaxRows = GrdCli.MaxRows - 1
    
    LblQtde.Caption = FormataNum(GrdCli.MaxRows)
    LblQtde.Visible = True
    RecResult.Close

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmConsultCli.hwnd)
    
    Height = 6120
    Width = 10425
    'Top = 1275
    'Left = 795

    Call MontaCbo

    LblQtde.Visible = False

    Screen.MousePointer = vbNormal
End Sub

Sub MontaCbo()
    Conecta
    
    Dim RecBai As New ADODB.Recordset
    Dim RecDt As New ADODB.Recordset
    
    StrSql = "Select distinct Bairro from tb_cliente order by Bairro asc"
    RecBai.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct DtCad from tb_cliente order by DtCad desc"
    RecDt.Open StrSql, vgCon, 1, 3
    
    Do While Not RecBai.EOF
        CboBairro.AddItem (RecBai.Fields.Item(0).Value)
        RecBai.MoveNext
    Loop
    
    Do While Not RecDt.EOF
        CboDtCad.AddItem (FormataData(RecDt.Fields.Item(0).Value))
        RecDt.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub Form_Resize()
  FrmConsultCli.Left = (MDIPrincipal.Width / 2) - (FrmConsultCli.Width / 1.93)
  FrmConsultCli.Top = (MDIPrincipal.Height / 3) - (FrmConsultCli.Height / 5)
End Sub

Private Sub GrdCli_DblClick(ByVal Col As Long, ByVal Row As Long)
    
    GrdCli.Row = Row
    GrdCli.Col = 1
        
    VPStrResponse = MsgBox("Deseja realmente excluir esse cliente?", vbYesNo)
    
    If VPStrResponse = vbYes Then
        Conecta
        
        Dim RecNumCart As New ADODB.Recordset
        
        StrSql = "Select NumCartao from tb_cartao where CodCli=" & GrdCli.Text
        RecNumCart.Open StrSql, vgCon, 1, 3
        
        vgCon.Execute ("Delete from tb_credito where NumCartao=" & RecNumCart.Fields.Item(0).Value)
        vgCon.Execute ("Delete from tb_guardacredito where NumCartao=" & RecNumCart.Fields.Item(0).Value)
        vgCon.Execute ("Delete from tb_cartao where CodCli=" & GrdCli.Text)
        vgCon.Execute ("Delete from tb_acesso where CodCli=" & GrdCli.Text)
        vgCon.Execute ("Delete from tb_cliente where CodCli=" & GrdCli.Text)
        
        Desconecta
        Me.CboBairro.Clear
        Me.CboDtCad.Clear
        Me.MontaCbo
        Me.CmdConsultar.Value = True
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
