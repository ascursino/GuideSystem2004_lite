VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmConsCaixa 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Movimento de Caixa"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10995
   Icon            =   "FrmConsCaixa.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   10995
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9480
      TabIndex        =   5
      ToolTipText     =   "Consulta movimento do caixa"
      Top             =   1440
      Width           =   855
   End
   Begin VB.Frame FraCaixa 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Caixa"
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   4695
      Begin VB.ComboBox CboDtItem2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsCaixa.frx":000C
         Left            =   2880
         List            =   "FrmConsCaixa.frx":000E
         TabIndex        =   1
         ToolTipText     =   "Maior data"
         Top             =   360
         Width           =   1455
      End
      Begin VB.ComboBox CboDtItem1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsCaixa.frx":0010
         Left            =   1080
         List            =   "FrmConsCaixa.frx":0012
         TabIndex        =   0
         ToolTipText     =   "Menor data"
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton CmdConsCx 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   1920
         TabIndex        =   2
         ToolTipText     =   "Consulta movimento do caixa"
         Top             =   840
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtItem 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmConsCaixa.frx":0014
         TabIndex        =   3
         Top             =   360
         Width           =   495
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "FrmConsCaixa.frx":007C
      Top             =   480
   End
   Begin FPSpread.vaSpread GrdCaixa 
      Height          =   3135
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   10530
      _Version        =   393216
      _ExtentX        =   18574
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
      SpreadDesigner  =   "FrmConsCaixa.frx":02B0
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmConsCaixa"
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
Public VPIntCred As Currency
Public VPIntDeb As Currency
Public VPIntTotal As Currency
Public VPStrResponse As String

Private Sub CmdConsCx_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select CodItem,Descr,Vldeb,Vlcred,DtItem from tb_caixa where 0=0"
        
     If CboDtItem1.Text <> "" And CboDtItem2.Text <> "" Then
         StrSql = StrSql + " and DtItem >=#" & FormataDataUS(CboDtItem1.Text) & "# and DtItem <=#" & FormataDataUS(CboDtItem2.Text) & "#"
     
     ElseIf CboDtItem1.Text = "" And CboDtItem2.Text = "" Then
         StrSql = StrSql + " and DtItem =#" & FormataDataUS(Date) & "#"
     
     ElseIf CboDtItem1.Text <> "" And CboDtItem2.Text = "" Then
         StrSql = StrSql + " and DtItem =#" & FormataDataUS(CboDtItem1.Text) & "#"
     
     ElseIf CboDtItem1.Text = "" And CboDtItem2.Text <> "" Then
         StrSql = StrSql + " and DtItem =#" & FormataDataUS(CboDtItem2.Text) & "#"
     
     End If
     
     StrSql = StrSql + " order by DtItem,CodItem asc"
     RecResult.Open StrSql, vgCon, 1, 3
        
     Call MontaGridCx
     
    Desconecta
    
    CmdImprimir.Enabled = True
    Screen.MousePointer = vbNormal

End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim Item As String
    Dim data As String
    Dim descr As String
    Dim cred As String
    Dim deb As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GrdCaixa.MaxRows
        
        GrdCaixa.Col = 1
        GrdCaixa.Row = VLStrLinha
        Item = GrdCaixa.Text
        
        GrdCaixa.Col = 2
        GrdCaixa.Row = VLStrLinha
        data = GrdCaixa.Text
        
        GrdCaixa.Col = 3
        GrdCaixa.Row = VLStrLinha
        descr = GrdCaixa.Text
        
        GrdCaixa.Col = 4
        GrdCaixa.Row = VLStrLinha
        cred = GrdCaixa.Text
        
        GrdCaixa.Col = 5
        GrdCaixa.Row = VLStrLinha
        deb = GrdCaixa.Text
        
        vgCon.Execute "INSERT INTO tb_auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & Item & "','" & data & "','" & descr & "','" & cred & "','" & deb & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCaixa.Show
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmConsCaixa.hwnd)
    
    Height = 5700
    Width = 11085
    'Top = 2070
    'Left = 1995
    
    Call MontaCbos
    
    CboDtItem1.Text = FormataData(Date)
    CboDtItem2.Text = FormataData(Date)
    
    CmdImprimir.Enabled = False
    
    Screen.MousePointer = vbNormal
    
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecData As New ADODB.Recordset
    
    StrSql = "Select distinct DtItem from tb_caixa order by DtItem desc"
    RecData.Open StrSql, vgCon, 1, 3
    
    Do While Not RecData.EOF
        CboDtItem1.AddItem (FormataData(RecData.Fields.Item(0).Value))
        CboDtItem2.AddItem (FormataData(RecData.Fields.Item(0).Value))
        RecData.MoveNext
    Loop
    
    Desconecta

End Sub

Sub MontaGridCx()
    
    If RecResult.EOF Then
           VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Guide System - Informação")
    End If
    
    VPIntLinha = 1
    VPIntCred = 0
    VPIntDeb = 0
    VPIntTotal = 0
    
    GrdCaixa.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        'VPIntCred = 0
        'VPIntDeb = 0
        'VPIntTotal = 0

        GrdCaixa.Row = VPIntLinha
        GrdCaixa.Lock = True
                        
        GrdCaixa.Col = 1   'CodItem
        GrdCaixa.Text = FormataNum(RecResult.Fields.Item(0).Value)
        GrdCaixa.Lock = True
        
        GrdCaixa.Col = 2   'Data
        GrdCaixa.Text = FormataData(RecResult.Fields.Item(4).Value)
        GrdCaixa.Lock = True
        
        GrdCaixa.Col = 3   'Descrição
        GrdCaixa.Text = RecResult.Fields.Item(1).Value
        GrdCaixa.Lock = True
        
        GrdCaixa.Col = 4   'Crédito
        If RecResult.Fields.Item(3).Value = "0" Then
            GrdCaixa.Text = ""
        Else
            GrdCaixa.Text = FormataMoeda(RecResult.Fields.Item(3).Value)
        End If
        GrdCaixa.Lock = True
           
        GrdCaixa.Col = 5   'Débito
        If RecResult.Fields.Item(2).Value = "0" Then
            GrdCaixa.Text = ""
        Else
            GrdCaixa.Text = FormataMoeda(RecResult.Fields.Item(2).Value)
        End If
        GrdCaixa.Lock = True
        
        VPIntCred = VPIntCred + RecResult.Fields.Item(3).Value
        VPIntDeb = VPIntDeb + RecResult.Fields.Item(2).Value
                
        VPIntLinha = VPIntLinha + 1
        
        GrdCaixa.MaxRows = GrdCaixa.MaxRows + 1
        RecResult.MoveNext
    Loop
    
    GrdCaixa.MaxRows = GrdCaixa.MaxRows + 2
    
    GrdCaixa.Row = GrdCaixa.MaxRows
    
    GrdCaixa.Col = 3   'Subtotal
    GrdCaixa.TypeHAlign = TypeHAlignRight
    GrdCaixa.FontBold = True
    GrdCaixa.Text = "SUBTOTAL:"
    GrdCaixa.Lock = True
    
    GrdCaixa.Col = 4   'Total de Crédito
    GrdCaixa.FontBold = True
    GrdCaixa.Text = FormataMoeda(VPIntCred)
    GrdCaixa.Lock = True
    
    GrdCaixa.Col = 5   'Total de Débito
    GrdCaixa.FontBold = True
    GrdCaixa.Text = FormataMoeda(VPIntDeb)
    GrdCaixa.Lock = True
    
    VPIntTotal = VPIntCred - VPIntDeb
    
    GrdCaixa.MaxRows = GrdCaixa.MaxRows + 2
    
    GrdCaixa.Row = GrdCaixa.MaxRows
    
    GrdCaixa.Col = 3   'Total
    GrdCaixa.TypeHAlign = TypeHAlignRight
    GrdCaixa.FontBold = True
    GrdCaixa.Text = "TOTAL:"
    GrdCaixa.Lock = True
    
    GrdCaixa.Col = 4   'Diferença entre Crédito e Débito
    If Mid(VPIntTotal, 1, 1) = "-" Then
        GrdCaixa.ForeColor = vbRed
    End If
    GrdCaixa.FontBold = True
    GrdCaixa.Text = FormataMoeda(VPIntTotal)
    GrdCaixa.Lock = True
    
    RecResult.Close
    
End Sub

Private Sub Form_Resize()
  FrmConsCaixa.Left = (MDIPrincipal.Width / 2) - (FrmConsCaixa.Width / 1.93)
  FrmConsCaixa.Top = (MDIPrincipal.Height / 3) - (FrmConsCaixa.Height / 5)
End Sub

Private Sub GrdCaixa_DblClick(ByVal Col As Long, ByVal Row As Long)
    GrdCaixa.Row = Row
    GrdCaixa.Col = 1
    
    If GrdCaixa.Text <> "" And GrdCaixa.Text <> "Cód. Item" Then
        VGIntCodItem = GrdCaixa.Text
            
        VPStrResponse = MsgBox("Deseja excluir este item?", vbYesNo)
        
        If VPStrResponse = vbYes Then
            Screen.MousePointer = vbHourglass
            
            VGStrLocalTemp = "conscaixa"
            FrmSenha.Show
        Else
            VGIntCodItem = 0
        End If
    End If
End Sub

Sub Exclui_Item()
        
    If VGStrSenha = "sim" Then
        Conecta
        
        Dim RecCx As New ADODB.Recordset
        
        StrSql = "Delete from tb_caixa where CodItem=" & VGIntCodItem
        RecCx.Open StrSql, vgCon, 1, 3
        
        Desconecta
        
        VGIntCodItem = 0
        
        Me.CmdConsCx.Value = True
    End If

End Sub
