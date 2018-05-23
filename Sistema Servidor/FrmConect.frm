VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmConect 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Usuários conectados"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7890
   Icon            =   "FrmConect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   7890
   ShowInTaskbar   =   0   'False
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   6480
      OleObjectBlob   =   "FrmConect.frx":000C
      Top             =   3360
   End
   Begin FPSpread.vaSpread GrdCon 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _Version        =   393216
      _ExtentX        =   13361
      _ExtentY        =   7223
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
      MaxCols         =   4
      MaxRows         =   0
      OperationMode   =   2
      RetainSelBlock  =   0   'False
      RowHeaderDisplay=   0
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      ShadowColor     =   12632256
      SpreadDesigner  =   "FrmConect.frx":0240
      UserResize      =   1
   End
End
Attribute VB_Name = "FrmConect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrResponse As String
Public VPStrNomCli As String
'Public VPIntNumCartTemp As Integer

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmConect.hwnd)
    
    Height = 4800
    Width = 8145
    'Top = 2070
    'Left = 1980
    
    Call MontaGridConect
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaGridConect()
    Conecta
    
    Dim RecCon As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    Dim RecCart As New ADODB.Recordset
        
    StrSql = "Select CodCli,NumMaq,NumCartao from tb_conect order by NumMaq"
    RecCon.Open StrSql, vgCon, 1, 3
        
    'If RecResult.EOF Then
    '       VPStrBox = MsgBox("Lista de espera vazia.", vbInformation, "Guide System - Informação")
    'End If
    
    'If IsNull(RecCon.Fields.Item(2).Value) = True Then
    '    VPIntNumCartTemp = 0
    'Else
    '    VPIntNumCartTemp = RecCon.Fields.Item(2).Value
    'End If
    
    VPIntLinha = 1
    
    GrdCon.MaxRows = VPIntLinha
           
    Do While Not RecCon.EOF
        
        StrSql = "Select Nome from tb_cliente where CodCli=" & RecCon.Fields.Item(0).Value
        RecCli.Open StrSql, vgCon, 1, 3
        
        StrSql = "Select TempoRest from tb_credito where NumCartao=" & RecCon.Fields.Item(2).Value
        RecCart.Open StrSql, vgCon, 1, 3
        
        If RecCli.EOF Then  'não achou nada
            VPStrNomCli = "CLIENTE EXCLUÍDO"
        Else
            VPStrNomCli = RecCli.Fields.Item(0).Value
        End If
        
        GrdCon.Row = VPIntLinha
        GrdCon.Lock = True
                        
        GrdCon.Col = 1  'Nº da máquina
        GrdCon.Text = FormataNum(RecCon.Fields.Item(1).Value)    'NumMaq
        GrdCon.Lock = True
        
        GrdCon.Col = 2  'Tempo Restante
        GrdCon.Text = RecCart.Fields.Item(0).Value
        GrdCon.Lock = True
        
        GrdCon.Col = 3  'Cód. Cliente
        GrdCon.Text = FormataNum(RecCon.Fields.Item(0).Value)
        GrdCon.Lock = True
        
        GrdCon.Col = 4  'Nome do Cliente
        GrdCon.Text = VPStrNomCli
        GrdCon.Lock = True
        
        VPIntLinha = VPIntLinha + 1
        
        GrdCon.MaxRows = GrdCon.MaxRows + 1
        RecCon.MoveNext
        RecCli.Close
    Loop
    GrdCon.MaxRows = GrdCon.MaxRows - 1
    RecCon.Close
    
    Desconecta
End Sub

Private Sub Form_Resize()
  FrmConect.Left = (MDIPrincipal.Width / 2) - (FrmConect.Width / 1.93)
  FrmConect.Top = (MDIPrincipal.Height / 3) - (FrmConect.Height / 5)
End Sub

Private Sub GrdCon_DblClick(ByVal Col As Long, ByVal Row As Long)

    VPStrResponse = MsgBox("Desconectar usuário?", vbYesNo)
    
    If VPStrResponse = vbYes Then
        Conecta
        
        Dim RecLista As New ADODB.Recordset
        Dim RecMaq As New ADODB.Recordset
        Dim RecMaqCli As New ADODB.Recordset
        
        GrdCon.Row = Row
        GrdCon.Col = 1
        
        StrSql = "Delete from tb_conect where NumMaq=" & GrdCon.Text
        RecLista.Open StrSql, vgCon, 1, 3
        
        StrSql = "Update tb_maquina set Situacao='livre' where NumMaq=" & GrdCon.Text
        RecMaq.Open StrSql, vgCon, 1, 3
        
        Desconecta
        
        Me.MontaGridConect

    End If

End Sub

