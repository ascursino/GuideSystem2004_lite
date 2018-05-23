VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmPreco 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Produtos & Preços"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "FrmPreco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   600
      OleObjectBlob   =   "FrmPreco.frx":000C
      Top             =   1800
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      ToolTipText     =   "Consulta movimento do caixa"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdExcluir 
      Caption         =   "Excluir"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      ToolTipText     =   "Exclui produto"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdAlterar 
      Caption         =   "Alterar"
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      ToolTipText     =   "Altera dados do produto"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton CmdIncluir 
      Caption         =   "Incluir"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      ToolTipText     =   "Inclui produto"
      Top             =   2760
      Width           =   1215
   End
   Begin FPSpread.vaSpread GrdPreco 
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5915
      _Version        =   393216
      _ExtentX        =   10433
      _ExtentY        =   4471
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
      MaxRows         =   1
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmPreco.frx":0240
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmPreco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPIntLinha As Integer

Private Sub CmdAlterar_Click()
    Screen.MousePointer = vbHourglass
    
    VGStrPreco = "alterar"
    FrmPrecoIncAlt.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdExcluir_Click()
    Screen.MousePointer = vbHourglass
    
    VPStrResponse = MsgBox("Deseja realmente excluir este produto?", vbYesNo)
    
    If VPStrResponse = vbYes Then
        
        If VGIntCodProd = 1 Then
            VPStrBox = MsgBox("Esse produto é necessário ao sistema." & Chr(13) & "Não poderá ser excluído.", vbInformation, "Guide System - Informação")
        Else
            Conecta
            
            Dim RecProd As New ADODB.Recordset
            
            StrSql = "Delete from tb_preco where CodProd=" & VGIntCodProd
            RecProd.Open StrSql, vgCon, 1, 3
            
            Desconecta
            
            VGIntCodProd = 0
            
            Me.MontaGridPreco
        End If
    Else
        VGIntCodProd = 0
    End If
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim codprod As String
    Dim prod As String
    Dim preco As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GrdPreco.MaxRows
        
        GrdPreco.Col = 1
        GrdPreco.Row = VLStrLinha
        codprod = GrdPreco.Text
        
        GrdPreco.Col = 2
        GrdPreco.Row = VLStrLinha
        prod = GrdPreco.Text
        
        GrdPreco.Col = 3
        GrdPreco.Row = VLStrLinha
        preco = GrdPreco.Text
        
        vgCon.Execute "INSERT INTO tb_auxiliar " & _
        "(campo01,campo02,campo03) " & _
        "VALUES ('" & codprod & "','" & prod & "','" & preco & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptPreco.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdIncluir_Click()
    Screen.MousePointer = vbHourglass
    
    VGStrPreco = "incluir"
    FrmPrecoIncAlt.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmPreco.hwnd)
    
    Height = 3735
    Width = 6585
    'Top = 1275
    'Left = 2640
    
    Call MontaGridPreco

    CmdAlterar.Enabled = False
    CmdExcluir.Enabled = False
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaGridPreco()
    
    Conecta
    
    Dim RecPr As New ADODB.Recordset
    
    StrSql = "Select CodProd,Prod,Preco from tb_preco order by Prod"
    RecPr.Open StrSql, vgCon, 1, 3
        
    'If RecPr.EOF Then
    '       VPStrBox = MsgBox("Lista sem produtos", vbInformation, "Guide System - Informação")
    'End If
   
    VPIntLinha = 1
    
    GrdPreco.MaxRows = VPIntLinha
           
    Do While Not RecPr.EOF
        
        GrdPreco.Row = VPIntLinha
        GrdPreco.Lock = True
                        
        GrdPreco.Col = 1
        GrdPreco.Text = FormataNum(RecPr.Fields.Item(0).Value)
        GrdPreco.Lock = True
        
        GrdPreco.Col = 2
        GrdPreco.Text = RecPr.Fields.Item(1).Value
        GrdPreco.Lock = True
        
        GrdPreco.Col = 3
        GrdPreco.Text = FormataMoeda(RecPr.Fields.Item(2).Value)
        GrdPreco.Lock = True
        
        
        VPIntLinha = VPIntLinha + 1
        
        GrdPreco.MaxRows = GrdPreco.MaxRows + 1
        RecPr.MoveNext
    Loop
    GrdPreco.MaxRows = GrdPreco.MaxRows - 1
    RecPr.Close
    
    Desconecta

End Sub

Private Sub Form_Resize()
  FrmPreco.Left = (MDIPrincipal.Width / 2) - (FrmPreco.Width / 1.93)
  FrmPreco.Top = (MDIPrincipal.Height / 3) - (FrmPreco.Height / 5)
End Sub

Private Sub GrdPreco_Click(ByVal Col As Long, ByVal Row As Long)
    GrdPreco.Row = Row
    GrdPreco.Col = 1
    
    If GrdPreco.Text <> "Cód.Prod" Then
        VGIntCodProd = GrdPreco.Text
        CmdAlterar.Enabled = True
        CmdExcluir.Enabled = True
    Else
        CmdAlterar.Enabled = False
        CmdExcluir.Enabled = False
    End If

End Sub
