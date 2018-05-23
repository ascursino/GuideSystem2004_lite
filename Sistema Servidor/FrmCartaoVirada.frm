VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmCartaoVirada 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Inserir Créditos em Cartões de Virada"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10065
   Icon            =   "FrmCartaoVirada.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5820
   ScaleWidth      =   10065
   Begin VB.Frame FraRecarga 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5655
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   9615
      Begin VB.TextBox TxtCli26 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   25
         ToolTipText     =   "Número do cartão"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox TxtPreco 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7800
         TabIndex        =   26
         ToolTipText     =   "Número do cartão"
         Top             =   4320
         Width           =   1095
      End
      Begin VB.CommandButton CmdGerar 
         Caption         =   "Gerar Cartão"
         Height          =   375
         Left            =   7680
         TabIndex        =   27
         ToolTipText     =   "Inclui cadastro do cartão"
         Top             =   5040
         Width           =   1335
      End
      Begin VB.TextBox TxtCli24 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   23
         ToolTipText     =   "Número do cartão"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtCli23 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   22
         ToolTipText     =   "Número do cartão"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtCli22 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   21
         ToolTipText     =   "Número do cartão"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtCli25 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   8400
         TabIndex        =   24
         ToolTipText     =   "Número do cartão"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtCli21 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   20
         ToolTipText     =   "Número do cartão"
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox TxtCli18 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   17
         ToolTipText     =   "Número do cartão"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtCli06 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         ToolTipText     =   "Número do cartão"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox TxtCli07 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         ToolTipText     =   "Número do cartão"
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox TxtCli13 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   12
         ToolTipText     =   "Número do cartão"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox TxtCli14 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   13
         ToolTipText     =   "Número do cartão"
         Top             =   4800
         Width           =   1095
      End
      Begin VB.TextBox TxtCli15 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   14
         ToolTipText     =   "Número do cartão"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtCli16 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   15
         ToolTipText     =   "Número do cartão"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtCli17 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   16
         ToolTipText     =   "Número do cartão"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtCli19 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   18
         ToolTipText     =   "Número do cartão"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox TxtCli20 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   19
         ToolTipText     =   "Número do cartão"
         Top             =   4080
         Width           =   1095
      End
      Begin VB.TextBox TxtCli10 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   9
         ToolTipText     =   "Número do cartão"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtCli12 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   11
         ToolTipText     =   "Número do cartão"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox TxtCli11 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   10
         ToolTipText     =   "Número do cartão"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtCli09 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   8
         ToolTipText     =   "Número do cartão"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtCli08 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3600
         TabIndex        =   7
         ToolTipText     =   "Número do cartão"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtCli05 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   4
         ToolTipText     =   "Número do cartão"
         Top             =   3360
         Width           =   1095
      End
      Begin VB.TextBox TxtCli04 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   3
         ToolTipText     =   "Número do cartão"
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox TxtCli03 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         ToolTipText     =   "Número do cartão"
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox TxtCli02 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         ToolTipText     =   "Número do cartão"
         Top             =   1200
         Width           =   1095
      End
      Begin VB.TextBox TxtCli01 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         ToolTipText     =   "Número do cartão"
         Top             =   480
         Width           =   1095
      End
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   6960
         OleObjectBlob   =   "FrmCartaoVirada.frx":000C
         Top             =   5160
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart01 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0240
         TabIndex        =   29
         Top             =   240
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli01 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":02B6
         TabIndex        =   30
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart02 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":032E
         TabIndex        =   31
         Top             =   960
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart03 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":03A4
         TabIndex        =   32
         Top             =   1680
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart04 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":041A
         TabIndex        =   33
         Top             =   2400
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart05 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0490
         TabIndex        =   34
         Top             =   3120
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart06 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0506
         TabIndex        =   35
         Top             =   3840
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart07 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":057C
         TabIndex        =   36
         Top             =   4560
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart08 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":05F2
         TabIndex        =   37
         Top             =   240
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart09 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":0668
         TabIndex        =   38
         Top             =   960
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart10 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":06DE
         TabIndex        =   39
         Top             =   1680
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart11 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":0754
         TabIndex        =   40
         Top             =   2400
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart12 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":07CA
         TabIndex        =   41
         Top             =   3120
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart13 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":0840
         TabIndex        =   42
         Top             =   3840
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart14 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":08B6
         TabIndex        =   43
         Top             =   4560
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart15 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":092C
         TabIndex        =   44
         Top             =   240
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart16 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":09A2
         TabIndex        =   45
         Top             =   960
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart17 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":0A18
         TabIndex        =   46
         Top             =   1680
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart18 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":0A8E
         TabIndex        =   47
         Top             =   2400
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart19 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":0B04
         TabIndex        =   48
         Top             =   3120
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart20 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":0B7A
         TabIndex        =   49
         Top             =   3840
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart21 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":0BF0
         TabIndex        =   50
         Top             =   4560
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart22 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":0C66
         TabIndex        =   51
         Top             =   240
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart23 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":0CDC
         TabIndex        =   52
         Top             =   960
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart24 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":0D52
         TabIndex        =   53
         Top             =   1680
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart25 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":0DC8
         TabIndex        =   54
         Top             =   2400
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli02 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0E3E
         TabIndex        =   55
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli03 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0EB6
         TabIndex        =   56
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli04 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0F2E
         TabIndex        =   57
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli05 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":0FA6
         TabIndex        =   58
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli06 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":101E
         TabIndex        =   59
         Top             =   4080
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli07 
         Height          =   255
         Left            =   120
         OleObjectBlob   =   "FrmCartaoVirada.frx":1096
         TabIndex        =   60
         Top             =   4800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli08 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":110E
         TabIndex        =   61
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli09 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":1186
         TabIndex        =   62
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli10 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":11FE
         TabIndex        =   63
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli11 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":1276
         TabIndex        =   64
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli12 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":12EE
         TabIndex        =   65
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli13 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":1366
         TabIndex        =   66
         Top             =   4080
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli14 
         Height          =   255
         Left            =   2520
         OleObjectBlob   =   "FrmCartaoVirada.frx":13DE
         TabIndex        =   67
         Top             =   4800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli15 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":1456
         TabIndex        =   68
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli16 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":14CE
         TabIndex        =   69
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli17 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":1546
         TabIndex        =   70
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli18 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":15BE
         TabIndex        =   71
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli19 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":1636
         TabIndex        =   72
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli20 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":16AE
         TabIndex        =   73
         Top             =   4080
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli21 
         Height          =   255
         Left            =   4920
         OleObjectBlob   =   "FrmCartaoVirada.frx":1726
         TabIndex        =   74
         Top             =   4800
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli22 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":179E
         TabIndex        =   75
         Top             =   480
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli23 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":1816
         TabIndex        =   76
         Top             =   1200
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli24 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":188E
         TabIndex        =   77
         Top             =   1920
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli25 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":1906
         TabIndex        =   78
         Top             =   2640
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCart26 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":197E
         TabIndex        =   79
         Top             =   3120
         Width           =   2175
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblCli26 
         Height          =   255
         Left            =   7320
         OleObjectBlob   =   "FrmCartaoVirada.frx":19F4
         TabIndex        =   80
         Top             =   3360
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblPreco 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmCartaoVirada.frx":1A6C
         TabIndex        =   81
         Top             =   3960
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FrmCartaoVirada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrDescr As String
Public VPStrValCred As String

Private Sub CmdGerar_Click()
    Screen.MousePointer = vbHourglass

    Conecta
        
    Dim RecCli As New ADODB.Recordset
    Dim RecCxa As New ADODB.Recordset
    
    If TxtCli01.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1001'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli01.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1001"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli02.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1002'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli02.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1002"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli03.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1003'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli03.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1003"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli04.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1004'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli04.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1004"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli05.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1005'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli05.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1005"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli06.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1006'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli06.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1006"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli07.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1007'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli07.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1007"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli08.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1008'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli08.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1008"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli09.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1009'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli09.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1009"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli10.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1010'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli10.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1010"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli11.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1011'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli11.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1011"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli12.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1012'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli12.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1012"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli13.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1013'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli13.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1013"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli14.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1014'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli14.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1014"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli15.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1015'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli15.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1015"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli16.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1016'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli16.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1016"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli17.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1017'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli17.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1017"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli18.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1018'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli18.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1018"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli19.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1019'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli19.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1019"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli20.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1020'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli20.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1020"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli21.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1021'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli21.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1021"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli22.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1022'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli22.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1022"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli23.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1023'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli23.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1023"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli24.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1024'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli24.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1024"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli25.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1025'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli25.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1025"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    If TxtCli26.Text <> "" Then
        StrSql = "Select * from tb_cartaovirada where NumCartao='V1026'"
        RecCli.Open StrSql, vgCon, 1, 3
        
        RecCli("CodCli") = TxtCli26.Text
        RecCli("DtCartao") = FormataDataUS(Date)
        RecCli.Update
        
        'insere item na tabela de caixa
        VPStrDescr = "Crédito para cartão de virada V1026"
        VPStrValCred = TxtPreco.Text
        
        StrSql = "Select * from tb_caixa"
        RecCxa.Open StrSql, vgCon, 1, 3
        
        RecCxa.AddNew
        RecCxa("Descr") = VPStrDescr
        RecCxa("Vldeb") = "0"
        RecCxa("Vlcred") = VPStrValCred
        RecCxa("DtItem") = FormataDataUS(Date)
        RecCxa.Update
    End If
    
    Desconecta
    
    Screen.MousePointer = vbNormal
    
    VPStrBox = MsgBox("Cartão(ões) gerado(s).", vbInformation, "Informação")

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmCartaoVirada.hWnd)
    
    Height = 6225
    Width = 10185
    Top = 1275
    Left = 720
    
    Screen.MousePointer = vbNormal
End Sub

Private Sub Form_Resize()
  FrmCartaoVirada.Left = (MDIPrincipal.Width / 2) - (FrmCartaoVirada.Width / 2)
  FrmCartaoVirada.Top = (MDIPrincipal.Height / 3) - (FrmCartaoVirada.Height / 3)
End Sub
