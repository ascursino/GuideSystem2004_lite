VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmConsultCred 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Arquivo de Crédito"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9555
   Icon            =   "FrmConsultCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   7200
      OleObjectBlob   =   "FrmConsultCred.frx":000C
      Top             =   360
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      ToolTipText     =   "Consulta movimento do caixa"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame FraCredito 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cartão e Crédito"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   960
      TabIndex        =   6
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox CboNumCart 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCred.frx":0240
         Left            =   240
         List            =   "FrmConsultCred.frx":0242
         TabIndex        =   0
         ToolTipText     =   "Número do cartão"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox CboDtCred2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCred.frx":0244
         Left            =   2280
         List            =   "FrmConsultCred.frx":0246
         TabIndex        =   2
         ToolTipText     =   "Maior data"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboCred 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCred.frx":0248
         Left            =   4320
         List            =   "FrmConsultCred.frx":024A
         TabIndex        =   3
         ToolTipText     =   "Crédito"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox CboDtCred1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCred.frx":024C
         Left            =   2280
         List            =   "FrmConsultCred.frx":024E
         TabIndex        =   1
         ToolTipText     =   "Menor data"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdConsCred 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   5880
         TabIndex        =   4
         ToolTipText     =   "Consulta cartão e crédito"
         Top             =   960
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumCart 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConsultCred.frx":0250
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCred 
         Height          =   255
         Left            =   2280
         OleObjectBlob   =   "FrmConsultCred.frx":02C8
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCred 
         Height          =   255
         Left            =   4320
         OleObjectBlob   =   "FrmConsultCred.frx":0346
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblConsulta 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmConsultCred.frx":03B4
      TabIndex        =   11
      Top             =   1920
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblResConsulta 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmConsultCred.frx":042C
      TabIndex        =   5
      Top             =   1920
      Width           =   5055
   End
   Begin FPSpread.vaSpread GrdCredito 
      Height          =   3135
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   9060
      _Version        =   393216
      _ExtentX        =   15981
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
      MaxCols         =   5
      MaxRows         =   0
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmConsultCred.frx":0496
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmConsultCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecResult As New ADODB.Recordset
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
    Skin1.ApplySkin (FrmConsultCred.hwnd)
    
    Height = 6090
    Width = 9645
    'Top = 1275
    'Left = 300
    
    Call MontaCbos

    LblResConsulta.Visible = False
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecCli As New ADODB.Recordset
    Dim RecCodCli As New ADODB.Recordset
    Dim RecCart As New ADODB.Recordset
    Dim RecCred As New ADODB.Recordset
    Dim RecData As New ADODB.Recordset
    
    StrSql = "Select distinct NumCartao from tb_guardacredito"
    RecCart.Open StrSql, vgCon, 1, 3
    
    Do While Not RecCart.EOF
        CboNumCart.AddItem (RecCart.Fields.Item(0).Value)
        RecCart.MoveNext
    Loop
    
    StrSql = "Select distinct TempoCred from tb_guardacredito"
    RecCred.Open StrSql, vgCon, 1, 3
    
    Do While Not RecCred.EOF
        CboCred.AddItem (RecCred.Fields.Item(0).Value)
        RecCred.MoveNext
    Loop
    
    StrSql = "Select distinct DtCred from tb_guardacredito"
    RecData.Open StrSql, vgCon, 1, 3
    
    Do While Not RecData.EOF
        CboDtCred1.AddItem (FormataData(RecData.Fields.Item(0).Value))
        CboDtCred2.AddItem (FormataData(RecData.Fields.Item(0).Value))
        RecData.MoveNext
    Loop
    
    Desconecta

End Sub

Sub MontaGridCart()
    If RecResult.EOF Then
           VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Guide System - Informação")
    End If
   
    Dim RecCodCli As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    
    VPIntLinha = 1
    
    GrdCredito.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        
        StrSql = "Select distinct CodCli,NumCartao from tb_cartao where NumCartao=" & RecResult.Fields.Item(0).Value
        RecCodCli.Open StrSql, vgCon, 1, 3
        
        StrSql = "Select distinct CodCli,Nome from tb_cliente where CodCli=" & RecCodCli.Fields.Item(0).Value
        RecCli.Open StrSql, vgCon, 1, 3
        
        GrdCredito.Row = VPIntLinha
        GrdCredito.Lock = True
                        
        GrdCredito.Col = 1   'CodCli
        GrdCredito.Text = FormataNum(RecCli.Fields.Item(0).Value)
        GrdCredito.Lock = True
        
        GrdCredito.Col = 2   'Nome
        GrdCredito.Text = RecCli.Fields.Item(1).Value
        GrdCredito.Lock = True
        
        GrdCredito.Col = 3   'Nº Cartão
        GrdCredito.Text = FormataNum(RecResult.Fields.Item(0).Value)
        GrdCredito.Lock = True
        
        GrdCredito.Col = 4   'Data Crédito
        GrdCredito.Text = FormataData(RecResult.Fields.Item(2).Value)
        GrdCredito.Lock = True
           
        GrdCredito.Col = 5   'Crédito
        GrdCredito.Text = RecResult.Fields.Item(1).Value
        GrdCredito.Lock = True
        
        VPIntLinha = VPIntLinha + 1
        
        GrdCredito.MaxRows = GrdCredito.MaxRows + 1
        RecResult.MoveNext
        RecCodCli.Close
        RecCli.Close
    Loop

    GrdCredito.MaxRows = GrdCredito.MaxRows - 1
    RecResult.Close

End Sub

Private Sub Form_Resize()
  FrmConsultCred.Left = (MDIPrincipal.Width / 2) - (FrmConsultCred.Width / 1.93)
  FrmConsultCred.Top = (MDIPrincipal.Height / 3) - (FrmConsultCred.Height / 5)
End Sub
