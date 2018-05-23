VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmJogos 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Manutenção de Jogos"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10245
   Icon            =   "FrmJogos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10245
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FraAcesso 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Inclusão de jogos"
      ForeColor       =   &H80000008&
      Height          =   5775
      Left            =   6480
      TabIndex        =   1
      Top             =   120
      Width           =   3375
      Begin VB.TextBox TxtFaixa 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   17
         ToolTipText     =   "Faixa etária do jogo (censura)"
         Top             =   4800
         Width           =   2895
      End
      Begin VB.TextBox TxtArq2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   15
         ToolTipText     =   "Nome de outro arquivo executável necessário"
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox TxtArq1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   13
         ToolTipText     =   "Nome do arquivo executável do jogo"
         Top             =   3000
         Width           =   2895
      End
      Begin VB.TextBox TxtCaminho 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   11
         Text            =   "C:\"
         ToolTipText     =   "caminho físico da instalação do jogo"
         Top             =   2400
         Width           =   2895
      End
      Begin VB.TextBox TxtParam 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   9
         ToolTipText     =   "Parâmetros necessários para o jogo"
         Top             =   4200
         Width           =   2895
      End
      Begin VB.TextBox TxtImagem 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   7
         ToolTipText     =   "Imagem do jogo"
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox TxtSigla 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   240
         MaxLength       =   10
         TabIndex        =   4
         ToolTipText     =   "Sigla do jogo"
         Top             =   600
         Width           =   2895
      End
      Begin VB.CommandButton CmdIncAlt 
         Caption         =   "Incluir"
         Height          =   375
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Inclui dados do jogo"
         Top             =   5280
         Width           =   1095
      End
      Begin VB.TextBox TxtJogo 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   10
         TabIndex        =   2
         ToolTipText     =   "Nome do jogo"
         Top             =   1200
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblLogin 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":000C
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblSenha 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":0086
         TabIndex        =   6
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":00FE
         TabIndex        =   8
         Top             =   1560
         Width           =   855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":016A
         TabIndex        =   10
         Top             =   3960
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":01F6
         TabIndex        =   12
         Top             =   2160
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":0294
         TabIndex        =   14
         Top             =   2760
         Width           =   2295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":032C
         TabIndex        =   16
         Top             =   3360
         Width           =   2415
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel6 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmJogos.frx":03BC
         TabIndex        =   18
         Top             =   4560
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   5400
      OleObjectBlob   =   "FrmJogos.frx":0434
      Top             =   4440
   End
   Begin FPSpread.vaSpread GrdJogo 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6225
      _Version        =   393216
      _ExtentX        =   10980
      _ExtentY        =   10186
      _StockProps     =   64
      AllowDragDrop   =   -1  'True
      AllowMultiBlocks=   -1  'True
      AllowUserFormulas=   -1  'True
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
      Protect         =   0   'False
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmJogos.frx":0668
   End
End
Attribute VB_Name = "FrmJogos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPIntLinha As Integer
Public VPStrCancel As String
Public VPStrCred As String
Public VPStrDtCred As String

Private Sub CmdConsCred_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select NumCartao,TempoCred,DtCred from tb_guardacredito where 0=0"
        
     If CboDtCred1.Text <> "" Or CboDtCred2.Text <> "" Then
        If CboDtCred1.Text = "" Then
            CboDtCred1.Text = FormataData(Date)
        End If
        
        If CboDtCred2.Text = "" Then
            CboDtCred2.Text = FormataData(Date)
        End If
        
        StrSql = StrSql + " and DtCred >=#" & FormataDataUS(CboDtCred1.Text) & "# and DtCred <=#" & FormataDataUS(CboDtCred2.Text) & "#"
     End If
            
     If CboCred.Text <> "" Then
        StrSql = StrSql + " and TempoCred='" & CboCred.Text & "'"
     End If
            
     If CboNumCart.Text <> "" Then
        StrSql = StrSql + " and NumCartao=" & CboNumCart.Text & ""
     End If
    
    StrSql = StrSql + " order by NumCartao"
    RecResult.Open StrSql, vgCon, 1, 3
    
     Call MontaGridCart
     
     LblResConsulta.Caption = "Arquivo de Crédito"
     LblResConsulta.Visible = True
     
    Desconecta
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim nome As String
    Dim cartao As String
    Dim datacredito As String
    Dim credito As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GrdCredito.MaxRows
        
        GrdCredito.Col = 2
        GrdCredito.Row = VLStrLinha
        nome = GrdCredito.Text
        
        GrdCredito.Col = 3
        GrdCredito.Row = VLStrLinha
        cartao = GrdCredito.Text
        
        GrdCredito.Col = 4
        GrdCredito.Row = VLStrLinha
        datacredito = GrdCredito.Text
        
        GrdCredito.Col = 5
        GrdCredito.Row = VLStrLinha
        credito = GrdCredito.Text
        
        vgCon.Execute "INSERT INTO tb_auxiliar " & _
        "(campo01,campo02,campo03,campo04) " & _
        "VALUES ('" & nome & "','" & cartao & "','" & datacredito & "','" & credito & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCredito.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmJogos.hwnd)
    
    Height = 6570
    Width = 10335
    'Top = 1275
    'Left = 300
    
    Call MontaGridJogo

    Screen.MousePointer = vbNormal
End Sub

Sub MontaGridJogo()
    Dim RecResult As New ADODB.Recordset
    
    Conecta
    
    StrSql = "Select * from tb_jogo order by Jogo"
    RecResult.Open StrSql, vgCon, 1, 3
    
    If Not RecResult.EOF Then
        VPIntLinha = 1
        
        GrdJogo.MaxRows = VPIntLinha
               
        Do While Not RecResult.EOF
            
            GrdJogo.Row = VPIntLinha
            GrdJogo.Lock = True
                            
            GrdJogo.Col = 1   'Sigla
            GrdJogo.Text = RecResult("Sigla")
            GrdJogo.Lock = True
            
            GrdJogo.Col = 2   'Jogo
            GrdJogo.Text = RecResult("Jogo")
            GrdJogo.Lock = True
            
            GrdJogo.Col = 3   'CodJogo
            GrdJogo.Text = Val(RecResult("CodJogo"))
            GrdJogo.Lock = True
            
            VPIntLinha = VPIntLinha + 1
            
            GrdJogo.MaxRows = GrdJogo.MaxRows + 1
            RecResult.MoveNext
        Loop
        
        GrdJogo.MaxRows = GrdJogo.MaxRows - 1
        RecResult.Close
        
    End If
        
    Desconecta
        
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmJogos.Left = (MDIPrincipal.Width / 2) - (FrmJogos.Width / 1.93)
  FrmJogos.Top = (MDIPrincipal.Height / 3) - (FrmJogos.Height / 5)
End Sub
