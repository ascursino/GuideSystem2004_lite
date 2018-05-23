VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmConsultCart 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Cartão e Crédito"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11235
   Icon            =   "FrmConsultCart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2400
      OleObjectBlob   =   "FrmConsultCart.frx":000C
      Top             =   3360
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   9840
      TabIndex        =   11
      ToolTipText     =   "Consulta movimento do caixa"
      Top             =   1800
      Width           =   855
   End
   Begin VB.Frame FraCredito 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Crédito"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5040
      TabIndex        =   10
      Top             =   120
      Width           =   5775
      Begin VB.ComboBox CboDtCred2 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCart.frx":0240
         Left            =   120
         List            =   "FrmConsultCart.frx":0242
         TabIndex        =   4
         ToolTipText     =   "Maior data"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboCred 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCart.frx":0244
         Left            =   1560
         List            =   "FrmConsultCart.frx":0246
         TabIndex        =   5
         ToolTipText     =   "Crédito"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox CboCredRest 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCart.frx":0248
         Left            =   3000
         List            =   "FrmConsultCart.frx":024A
         TabIndex        =   6
         ToolTipText     =   "Créditos restantes"
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox CboDtCred1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCart.frx":024C
         Left            =   120
         List            =   "FrmConsultCart.frx":024E
         TabIndex        =   3
         ToolTipText     =   "Menor data"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdConsCred 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   4560
         TabIndex        =   7
         ToolTipText     =   "Consulta cartão e crédito"
         Top             =   1080
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblDtCred 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmConsultCart.frx":0250
         TabIndex        =   14
         Top             =   360
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCred 
         Height          =   255
         Left            =   1560
         OleObjectBlob   =   "FrmConsultCart.frx":02CE
         TabIndex        =   15
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCredRest 
         Height          =   255
         Left            =   3000
         OleObjectBlob   =   "FrmConsultCart.frx":033C
         TabIndex        =   16
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame FraCartao 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Cartão e Cliente"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4815
      Begin VB.ComboBox CboNumCart 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCart.frx":03BC
         Left            =   1920
         List            =   "FrmConsultCart.frx":03BE
         TabIndex        =   1
         ToolTipText     =   "Número do cartão"
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox CboCodCli 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmConsultCart.frx":03C0
         Left            =   240
         List            =   "FrmConsultCart.frx":03C2
         TabIndex        =   0
         ToolTipText     =   "Código do cliente"
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton CmdConsCart 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   3600
         TabIndex        =   2
         ToolTipText     =   "Consulta cartão e crédito"
         Top             =   720
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   240
         OleObjectBlob   =   "FrmConsultCart.frx":03C4
         TabIndex        =   12
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumCart 
         Height          =   255
         Left            =   1920
         OleObjectBlob   =   "FrmConsultCart.frx":043C
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblConsulta 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmConsultCart.frx":04B4
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblResConsulta 
      Height          =   255
      Left            =   1200
      OleObjectBlob   =   "FrmConsultCart.frx":052C
      TabIndex        =   18
      Top             =   1920
      Width           =   2655
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblTotalCredRest 
      Height          =   255
      Left            =   4680
      OleObjectBlob   =   "FrmConsultCart.frx":0596
      TabIndex        =   19
      Top             =   1920
      Width           =   1935
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblVlCredRest 
      Height          =   255
      Left            =   6720
      OleObjectBlob   =   "FrmConsultCart.frx":062C
      TabIndex        =   8
      Top             =   1920
      Width           =   1815
   End
   Begin FPSpread.vaSpread GrdCartao 
      Height          =   3015
      Left            =   120
      TabIndex        =   20
      Top             =   2280
      Width           =   10695
      _Version        =   393216
      _ExtentX        =   18865
      _ExtentY        =   5318
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
      MaxCols         =   10
      MaxRows         =   0
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmConsultCart.frx":0690
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmConsultCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecResult As New ADODB.Recordset
Public VPStrBox As String
Public VPIntLinha As Integer
Public VPStrCancel As String
Public VPStrConsulta As String
Public VPStrCred As String
Public VPStrCredRest As String
Public VPStrDtCred As String
Public VPIntHora As Integer
Public VPIntMin As Integer
Public VPIntSeg As Integer
Public VPIntHoraRest As Integer
Public VPIntMinRest As Integer
Public VPIntSegRest As Integer
Public VPIntRest As Integer

Private Sub CmdConsCart_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select * from tb_cartao where 0=0"
        
     If CboNumCart.Text <> "" Then
        StrSql = StrSql + " and NumCartao=" & CboNumCart.Text & ""
     End If
     
     If CboCodCli.Text <> "" Then
        StrSql = StrSql + " and CodCli=" & CboCodCli.Text & ""
     End If
       
     If CboCodCli.Text <> "" Then
        StrSql = StrSql + " order by CodCli,NumCartao"
     
     ElseIf CboNumCart.Text <> "" Then
        StrSql = StrSql + " order by NumCartao"
     
     Else
        StrSql = StrSql + " order by CodCli,NumCartao"
     
     End If
     
     RecResult.Open StrSql, vgCon, 1, 3
     
     VPStrConsulta = "Cartao"
     
     Call MontaGridCart
     
     LblResConsulta.Caption = "Cartão e Cliente"
     LblResConsulta.Visible = True
     Desconecta
     
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdConsCred_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select * from tb_credito where 0=0"
        
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
            
     If CboCredRest.Text <> "" Then
        StrSql = StrSql + " and TempoRest='" & CboCredRest.Text & "'"
     End If
    
     StrSql = StrSql + " order by NumCartao"
    
     RecResult.Open StrSql, vgCon, 1, 3
        
     VPStrConsulta = "Credito"
     
     Call MontaGridCart
     
     LblResConsulta.Caption = "Crédito"
     LblResConsulta.Visible = True
     
    Desconecta
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim codigo As String
    Dim cartao As String
    Dim creditorestante As String
    Dim datacredito As String
    Dim cancelado As String
    Dim motivo As String
    Dim resp As String
    Dim datacancel As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GrdCartao.MaxRows
        
        GrdCartao.Col = 1
        GrdCartao.Row = VLStrLinha
        codigo = GrdCartao.Text
        
        GrdCartao.Col = 2
        GrdCartao.Row = VLStrLinha
        cartao = GrdCartao.Text
        
        GrdCartao.Col = 5
        GrdCartao.Row = VLStrLinha
        creditorestante = GrdCartao.Text
        
        GrdCartao.Col = 6
        GrdCartao.Row = VLStrLinha
        datacredito = GrdCartao.Text
        
        GrdCartao.Col = 7
        GrdCartao.Row = VLStrLinha
        cancelado = GrdCartao.Text
        
        GrdCartao.Col = 8
        GrdCartao.Row = VLStrLinha
        motivo = GrdCartao.Text
        
        GrdCartao.Col = 9
        GrdCartao.Row = VLStrLinha
        resp = GrdCartao.Text
        
        GrdCartao.Col = 10
        GrdCartao.Row = VLStrLinha
        datacancel = GrdCartao.Text
        
        vgCon.Execute "INSERT INTO tb_auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08) " & _
        "VALUES ('" & codigo & "','" & cartao & "','" & creditorestante & "','" & datacredito & "','" & cancelado & "','" & motivo & "','" & resp & "','" & datacancel & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCartao.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmConsultCart.hwnd)
    
    Height = 5880
    Width = 11325
    'Top = 2070
    'Left = 90
    
    Call MontaCbos

    LblTotalCredRest.Visible = False
    LblVlCredRest.Visible = False

    LblResConsulta.Visible = False

    If VGStrConsultCart = "sim" Then
        CboCodCli.Text = FormataNum(VGIntCodCli)

        VGIntCodCli = 0
        VGStrConsultCart = ""

        CmdConsCart.Value = True

    End If
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecCli As New ADODB.Recordset
    Dim RecCart As New ADODB.Recordset
    Dim RecCred As New ADODB.Recordset
    Dim RecData As New ADODB.Recordset
    
    StrSql = "Select distinct CodCli from tb_cliente"
    RecCli.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct NumCartao from tb_cartao"
    RecCart.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct TempoCred,TempoRest from tb_credito"
    RecCred.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct DtCred from tb_credito"
    RecData.Open StrSql, vgCon, 1, 3
    
    Do While Not RecCli.EOF
        CboCodCli.AddItem (RecCli.Fields.Item(0).Value)
        RecCli.MoveNext
    Loop
    
    Do While Not RecCart.EOF
        CboNumCart.AddItem (RecCart.Fields.Item(0).Value)
        RecCart.MoveNext
    Loop
    
    Do While Not RecCred.EOF
        CboCred.AddItem (RecCred.Fields.Item(0).Value)
        CboCredRest.AddItem (RecCred.Fields.Item(1).Value)
        RecCred.MoveNext
    Loop
    
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
    
   VPStrRest = "00:00:00"
    
    If VPStrConsulta = "Cartao" Then
       
        Dim RecCred As New ADODB.Recordset
        
        VPIntLinha = 1
        
        GrdCartao.MaxRows = VPIntLinha
               
        Do While Not RecResult.EOF
    
            StrSql = "Select TempoCred,TempoRest,DtCred from tb_credito where NumCartao=" & RecResult.Fields.Item(0).Value
            RecCred.Open StrSql, vgCon, 1, 3
            
            If RecCred.EOF Then
                VPStrCred = ""
                VPStrCredRest = ""
                VPStrDtCred = ""
            Else
                VPStrCred = RecCred.Fields.Item(0).Value
                VPStrCredRest = RecCred.Fields.Item(1).Value
                VPStrDtCred = FormataData(RecCred.Fields.Item(2).Value)
            End If
            
            GrdCartao.Row = VPIntLinha
            GrdCartao.Lock = True
                            
            GrdCartao.Col = 1   'CodCli
            GrdCartao.Text = FormataNum(RecResult.Fields.Item(1).Value)
            GrdCartao.Lock = True
            
            GrdCartao.Col = 2   'Nº Cartão
            GrdCartao.Text = FormataNum(RecResult.Fields.Item(0).Value)
            GrdCartao.Lock = True
            
            GrdCartao.Col = 3   'Data Cartão
            GrdCartao.Text = FormataData(RecResult.Fields.Item(2).Value)
            GrdCartao.Lock = True
               
            GrdCartao.Col = 4   'Crédito
            GrdCartao.Text = VPStrCred
            GrdCartao.Lock = True
               
            GrdCartao.Col = 5   'Crédito Restante
            GrdCartao.Text = VPStrCredRest
            GrdCartao.Lock = True
            
            'VPIntHora = Mid(VPStrCredRest, 1, 2)
            'VPIntMin = Mid(VPStrCredRest, 4, 2)
            'VPIntSeg = Mid(VPStrCredRest, 7, 2)
            
            'VPIntHoraRest = Mid(VPStrRest, 1, 2)
            'VPIntMinRest = Mid(VPStrRest, 4, 2)
            'VPIntSegRest = Mid(VPStrRest, 7, 2)
            
            'VPStrRest = FormataNum(VPIntHoraRest + VPIntHora) & ":" & FormataNum(VPIntMinRest + VPIntMin) & ":" & FormataNum(VPIntSegRest + VPIntSeg)
            
            GrdCartao.Col = 6   'Data Crédito
            GrdCartao.Text = VPStrDtCred
            GrdCartao.Lock = True
            
            If RecResult.Fields.Item(3).Value = True Then
                VPStrCancel = "Sim"
            Else
                VPStrCancel = "Não"
            End If
            
            GrdCartao.Col = 7   'Cancelado
            GrdCartao.Text = VPStrCancel
            GrdCartao.Lock = True
            
            If IsNull(RecResult.Fields.Item(4).Value) Then
                VPStrMotivo = ""
            Else
                VPStrMotivo = RecResult.Fields.Item(4).Value
            End If
            
            GrdCartao.Col = 8   'Motivo
            GrdCartao.Text = VPStrMotivo
            GrdCartao.Lock = True
            
            If IsNull(RecResult.Fields.Item(5).Value) Then
                VPStrResp = ""
            Else
                VPStrResp = RecResult.Fields.Item(5).Value
            End If
            
            GrdCartao.Col = 9   'Responsável
            GrdCartao.Text = VPStrResp
            GrdCartao.Lock = True
            
            If IsNull(RecResult.Fields.Item(6).Value) Then
                VPStrDataCancel = ""
            Else
                VPStrDataCancel = FormataData(RecResult.Fields.Item(6).Value)
            End If
            
            GrdCartao.Col = 10  'Data Cancel.
            GrdCartao.Text = VPStrDataCancel
            GrdCartao.Lock = True
            
            VPIntLinha = VPIntLinha + 1
            
            GrdCartao.MaxRows = GrdCartao.MaxRows + 1
            RecResult.MoveNext
            RecCred.Close
        Loop
        
    ElseIf VPStrConsulta = "Credito" Then
    
        Dim RecCart As New ADODB.Recordset
        
        VPIntLinha = 1
        
        GrdCartao.MaxRows = VPIntLinha
               
        Do While Not RecResult.EOF
            
            StrSql = "Select NumCartao,CodCli,DtCartao,Cancelado,Motivo,Resp,DtCancel from tb_cartao where NumCartao=" & RecResult.Fields.Item(0).Value
            RecCart.Open StrSql, vgCon, 1, 3
            
            GrdCartao.Row = VPIntLinha
            GrdCartao.Lock = True
                            
            GrdCartao.Col = 1   'CodCli
            GrdCartao.Text = FormataNum(RecCart.Fields.Item(1).Value)
            GrdCartao.Lock = True
            
            GrdCartao.Col = 2   'Nº Cartão
            GrdCartao.Text = FormataNum(RecCart.Fields.Item(0).Value)
            GrdCartao.Lock = True
            
            GrdCartao.Col = 3   'Data Cartão
            GrdCartao.Text = FormataData(RecCart.Fields.Item(2).Value)
            GrdCartao.Lock = True
               
            GrdCartao.Col = 4   'Crédito
            GrdCartao.Text = RecResult.Fields.Item(1).Value
            GrdCartao.Lock = True
               
            GrdCartao.Col = 5   'Crédito Restante
            GrdCartao.Text = RecResult.Fields.Item(2).Value
            GrdCartao.Lock = True
               
            GrdCartao.Col = 6   'Data Crédito
            GrdCartao.Text = FormataData(RecResult.Fields.Item(3).Value)
            GrdCartao.Lock = True
            
            If RecCart.Fields.Item(3).Value = True Then
                VPStrCancel = "Sim"
            Else
                VPStrCancel = "Não"
            End If
            
            GrdCartao.Col = 7   'Cancelado
            GrdCartao.Text = VPStrCancel
            GrdCartao.Lock = True
            
            If IsNull(RecCart.Fields.Item(4).Value) Then
                VPStrMotivo = ""
            Else
                VPStrMotivo = RecCart.Fields.Item(4).Value
            End If
            
            GrdCartao.Col = 8   'Motivo
            GrdCartao.Text = VPStrMotivo
            GrdCartao.Lock = True
            
            If IsNull(RecCart.Fields.Item(5).Value) Then
                VPStrResp = ""
            Else
                VPStrResp = RecCart.Fields.Item(5).Value
            End If
            
            GrdCartao.Col = 9   'Responsável
            GrdCartao.Text = VPStrResp
            GrdCartao.Lock = True
            
            If IsNull(RecCart.Fields.Item(6).Value) Then
                VPStrDataCancel = ""
            Else
                VPStrDataCancel = FormataData(RecCart.Fields.Item(6).Value)
            End If
            
            GrdCartao.Col = 10  'Data Cancel.
            GrdCartao.Text = VPStrDataCancel
            GrdCartao.Lock = True
            
            VPIntLinha = VPIntLinha + 1
            
            GrdCartao.MaxRows = GrdCartao.MaxRows + 1
            RecResult.MoveNext
            RecCart.Close
        Loop
    
    End If
    GrdCartao.MaxRows = GrdCartao.MaxRows - 1
    RecResult.Close
    
    'LblVlCredRest.Caption = FormataHora(VPStrRest)

End Sub

Private Sub Form_Resize()
  FrmConsultCart.Left = (MDIPrincipal.Width / 2) - (FrmConsultCart.Width / 1.93)
  FrmConsultCart.Top = (MDIPrincipal.Height / 3) - (FrmConsultCart.Height / 5)
End Sub
