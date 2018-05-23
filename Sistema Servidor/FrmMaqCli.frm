VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmMaqCli 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Arquivo de Uso das Máquinas"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9120
   Icon            =   "FrmMaqCli.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   9120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   7800
      TabIndex        =   7
      ToolTipText     =   "Consulta movimento do caixa"
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame FraMaqCli 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Máquinas e Clientes"
      ForeColor       =   &H80000008&
      Height          =   2055
      Left            =   2280
      TabIndex        =   6
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox CboDataEntr 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmMaqCli.frx":000C
         Left            =   1200
         List            =   "FrmMaqCli.frx":000E
         TabIndex        =   2
         ToolTipText     =   "Data de entrada"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboDataSai 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmMaqCli.frx":0010
         Left            =   2760
         List            =   "FrmMaqCli.frx":0012
         TabIndex        =   3
         ToolTipText     =   "Data de saída"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboCodCli 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmMaqCli.frx":0014
         Left            =   1200
         List            =   "FrmMaqCli.frx":0016
         TabIndex        =   0
         ToolTipText     =   "Código do cliente"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox CboNumMaq 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmMaqCli.frx":0018
         Left            =   1200
         List            =   "FrmMaqCli.frx":001A
         TabIndex        =   1
         ToolTipText     =   "Número da máquina"
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton CmdConsMaq 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "Consulta de utilização das máquinas"
         Top             =   1560
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCodCli 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMaqCli.frx":001C
         TabIndex        =   8
         Top             =   360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumMaq 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMaqCli.frx":0094
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblData 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmMaqCli.frx":0108
         TabIndex        =   5
         Top             =   1080
         Width           =   975
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   1080
      OleObjectBlob   =   "FrmMaqCli.frx":0170
      Top             =   600
   End
   Begin FPSpread.vaSpread GrdMaq 
      Height          =   3015
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   8775
      _Version        =   393216
      _ExtentX        =   15478
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
      MaxCols         =   5
      MaxRows         =   0
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmMaqCli.frx":03A4
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmMaqCli"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecResult As New ADODB.Recordset
Public VPStrBox As String
Public VPIntLinha As Integer

Private Sub CmdConsMaq_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    StrSql = "Select CodCli,NumMaq,DataEntr,HoraEntr,DataSaida,HoraSaida from tb_maqcli where 0=0"
        
     If CboCodCli.Text <> "" Then
        StrSql = StrSql + " and CodCli='" & CboCodCli.Text & "'"
     End If
            
     If CboNumMaq.Text <> "" Then
        StrSql = StrSql + " and NumMaq=" & CboNumMaq.Text & ""
     End If
        
     If CboDataEntr.Text <> "" Or CboDataSai.Text <> "" Then
        If CboDataEntr.Text = "" Then
            CboDataEntr.Text = FormataData(Date)
        End If
        
        If CboDataSai.Text = "" Then
            CboDataSai.Text = FormataData(Date)
        End If
        
        StrSql = StrSql + " and DataEntr >='" & FormataDataUS(CboDataEntr.Text) & "' and Datasaida <='" & FormataDataUS(CboDataSai.Text) & "'"
     End If
    
     StrSql = StrSql + " order by NumMaq"
     
     RecResult.Open StrSql, vgCon, 1, 3
     
     Call MontaGridCart
     
    Desconecta
    
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub CmdImprimir_Click()
    Screen.MousePointer = vbHourglass
    
    Dim nome As String
    Dim maq As String
    Dim entrada As String
    Dim saida As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GrdMaq.MaxRows
        
        GrdMaq.Col = 2
        GrdMaq.Row = VLStrLinha
        nome = GrdMaq.Text
        
        GrdMaq.Col = 3
        GrdMaq.Row = VLStrLinha
        maq = GrdMaq.Text
        
        GrdMaq.Col = 4
        GrdMaq.Row = VLStrLinha
        entrada = GrdMaq.Text
        
        GrdMaq.Col = 5
        GrdMaq.Row = VLStrLinha
        saida = GrdMaq.Text
        
        vgCon.Execute "INSERT INTO tb_auxiliar " & _
        "(campo01,campo02,campo03,campo04) " & _
        "VALUES ('" & nome & "','" & maq & "','" & entrada & "','" & saida & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptMaquina.Show
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmMaqCli.hwnd)
    
    Height = 6120
    Width = 9360
    'Top = 1275
    'Left = 1125
    
    Call MontaCbos
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaCbos()
    Conecta
    
    Dim RecCli As New ADODB.Recordset
    Dim RecMaq As New ADODB.Recordset
    Dim RecDataE As New ADODB.Recordset
    Dim RecDataS As New ADODB.Recordset
    
    StrSql = "Select distinct CodCli from tb_maqcli"
    RecCli.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct NumMaq from tb_maqcli"
    RecMaq.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct DataEntr from tb_maqcli"
    RecDataE.Open StrSql, vgCon, 1, 3
    
    StrSql = "Select distinct DataSaida from tb_maqcli"
    RecDataS.Open StrSql, vgCon, 1, 3
    
    Do While Not RecCli.EOF
        CboCodCli.AddItem (RecCli.Fields.Item(0).Value)
        RecCli.MoveNext
    Loop
        
    Do While Not RecMaq.EOF
        CboNumMaq.AddItem (RecMaq.Fields.Item(0).Value)
        RecMaq.MoveNext
    Loop
        
    Do While Not RecDataE.EOF
        CboDataEntr.AddItem (RecDataE.Fields.Item(0).Value)
        RecDataE.MoveNext
    Loop

    Do While Not RecDataS.EOF
        If Not IsNull(RecDataS.Fields.Item(0).Value) Then
            CboDataSai.AddItem (RecDataS.Fields.Item(0).Value)
        End If

        RecDataS.MoveNext
    Loop
    
    RecCli.Close
    RecMaq.Close
    RecDataE.Close
    RecDataS.Close

    Desconecta
    

End Sub

Sub MontaGridCart()
    If RecResult.EOF Then
           VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Guide System - Informação")
    End If
   
    Dim RecCli As New ADODB.Recordset
    Dim VPStrEntr As String
    Dim VPStrSai As String
    Dim VPStrCli As String
    
    VPIntLinha = 1
    
    GrdMaq.MaxRows = VPIntLinha
           
    Do While Not RecResult.EOF
        
        StrSql = "Select Nome from tb_cliente where CodCli=" & RecResult.Fields.Item(0).Value
        RecCli.Open StrSql, vgCon, 1, 3
        
        If IsNull(RecResult.Fields.Item(2).Value) And IsNull(RecResult.Fields.Item(3).Value) Then
            VPStrEntr = ""
        Else
            VPStrEntr = FormataData(RecResult.Fields.Item(2).Value) & "   " & FormataHora(RecResult.Fields.Item(3).Value)
        End If
        
        If IsNull(RecResult.Fields.Item(4).Value) And IsNull(RecResult.Fields.Item(5).Value) Then
            VPStrSai = ""
        Else
            VPStrSai = FormataData(RecResult.Fields.Item(4).Value) & "   " & FormataHora(RecResult.Fields.Item(5).Value)
        End If
        
        If RecCli.EOF Then
            VPStrCli = "CLIENTE EXCLUÍDO"
        Else
            VPStrCli = RecCli.Fields.Item(0).Value
        End If
        
        GrdMaq.Row = VPIntLinha
        GrdMaq.Lock = True
                        
        GrdMaq.Col = 1   'CodCli
        GrdMaq.Text = FormataNum(RecResult.Fields.Item(0).Value)
        GrdMaq.Lock = True
        
        GrdMaq.Col = 2   'Nome
        GrdMaq.Text = VPStrCli
        GrdMaq.Lock = True
        
        GrdMaq.Col = 3   'NumMaq
        GrdMaq.Text = FormataNum(RecResult.Fields.Item(1).Value)
        GrdMaq.Lock = True
        
        GrdMaq.Col = 4   'Data / Hora de entrada
        GrdMaq.Text = VPStrEntr
        GrdMaq.Lock = True
        
        GrdMaq.Col = 5   'Data / Hora de saída
        GrdMaq.Text = VPStrSai
        GrdMaq.Lock = True
        
        VPIntLinha = VPIntLinha + 1
        
        GrdMaq.MaxRows = GrdMaq.MaxRows + 1
        RecResult.MoveNext
        RecCli.Close
    Loop

    GrdMaq.MaxRows = GrdMaq.MaxRows - 1
    RecResult.Close

End Sub

Private Sub Form_Resize()
  FrmMaqCli.Left = (MDIPrincipal.Width / 2) - (FrmMaqCli.Width / 1.93)
  FrmMaqCli.Top = (MDIPrincipal.Height / 3) - (FrmMaqCli.Height / 5)
End Sub
