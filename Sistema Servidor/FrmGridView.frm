VERSION 5.00
Object = "{B02F3647-766B-11CE-AF28-C3A2FBE76A13}#3.0#0"; "SPR32X30.ocx"
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FrmGridView 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   7260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8145
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7260
   ScaleWidth      =   8145
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   840
      OleObjectBlob   =   "FrmGridView.frx":0000
      Top             =   2640
   End
   Begin VB.CommandButton CmdFechar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Fecha janela"
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton CmdImprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Imprimi visualização"
      Top             =   6720
      Width           =   1095
   End
   Begin FPSpread.vaSpreadPreview VwGrid 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   7695
      _Version        =   196608
      _ExtentX        =   13573
      _ExtentY        =   11245
      _StockProps     =   96
      BorderStyle     =   1
      AllowUserZoom   =   -1  'True
      GrayAreaColor   =   12632256
      GrayAreaMarginH =   720
      GrayAreaMarginType=   0
      GrayAreaMarginV =   720
      PageBorderColor =   8388608
      PageBorderWidth =   2
      PageShadowColor =   0
      PageShadowWidth =   2
      PageViewPercentage=   100
      PageViewType    =   0
      ScrollBarH      =   1
      ScrollBarV      =   1
      ScrollIncH      =   360
      ScrollIncV      =   360
      PageMultiCntH   =   1
      PageMultiCntV   =   1
      PageGutterH     =   -1
      PageGutterV     =   -1
      OLEDropMode     =   1
   End
End
Attribute VB_Name = "FrmGridView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdFechar_Click()
    Unload Me
    'Unload FrmRelConf
    
    If VGStrImprimir = "caixa" Then
        FrmConsCaixa.Enabled = True
    
    ElseIf VGStrImprimir = "cartcred" Then
        FrmConsultCart.Enabled = True
    
    ElseIf VGStrImprimir = "cli" Then
        FrmConsultCli.Enabled = True
    
    ElseIf VGStrImprimir = "cred" Then
        FrmConsultCred.Enabled = True
    
    ElseIf VGStrImprimir = "maqcli" Then
        FrmMaqCli.Enabled = True
    
    ElseIf VGStrImprimir = "prod" Then
        FrmPreco.Enabled = True
    
    End If
    
End Sub

Private Sub CmdImprimir_Click()
    
    If VGStrImprimir = "caixa" Then
        FrmConsCaixa.GrdCaixa.Action = ActionPrint
    
    ElseIf VGStrImprimir = "cartcred" Then
        FrmConsultCart.GrdCartao.Action = ActionPrint
        
    ElseIf VGStrImprimir = "cli" Then
        FrmConsultCli.GrdCli.Action = ActionPrint
        
    ElseIf VGStrImprimir = "cred" Then
        FrmConsultCred.GrdCredito.Action = ActionPrint
        
    ElseIf VGStrImprimir = "maqcli" Then
        FrmMaqCli.GrdMaq.Action = ActionPrint
        
    ElseIf VGStrImprimir = "prod" Then
        FrmPreco.GrdPreco.Action = ActionPrint
        
    End If

End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\Zhelezo.skn")
    Skin1.ApplySkin (FrmGridView.hWnd)
    
    Top = 60
    Left = 1740
    Height = 7665
    Width = 8265
    
    FrmGridView.Caption = font3 & " (Visualização)"
        
''    If VGStrImprimir = "caixa" Then
''        FrmConsCaixa.Enabled = False
''        font5 = "Total de " & FrmConsCaixa.GrdCaixa.PrintPageCount & " página(s)"
''        font6 = FormataData(Date) & "   " & Time
''        FrmConsCaixa.GrdCaixa.PrintFooter = "/r" & font6 & "/n/n/r" & font5
''        VwGrid.hWndSpread = FrmConsCaixa.GrdCaixa.hWnd
''
''    ElseIf VGStrImprimir = "cartcred" Then
''        FrmConsultCart.Enabled = False
''        font5 = "Total de " & FrmConsultCart.GrdCartao.PrintPageCount & " página(s)"
''        font6 = FormataData(Date) & "   " & Time
''        FrmConsultCart.GrdCartao.PrintFooter = "/r" & font6 & "/n/n/r" & font5
''        VwGrid.hWndSpread = FrmConsultCart.GrdCartao.hWnd
''
''    ElseIf VGStrImprimir = "cli" Then
''        FrmConsultCli.Enabled = False
''        font5 = "Total de " & FrmConsultCli.GrdCli.PrintPageCount & " página(s)"
''        font6 = FormataData(Date) & "   " & Time
''        FrmConsultCli.GrdCli.PrintFooter = "/r" & font6 & "/n/n/r" & font5
''        VwGrid.hWndSpread = FrmConsultCli.GrdCli.hWnd
''
''    ElseIf VGStrImprimir = "cred" Then
''        FrmConsultCred.Enabled = False
''        font5 = "Total de " & FrmConsultCred.GrdCredito.PrintPageCount & " página(s)"
''        font6 = FormataData(Date) & "   " & Time
''        FrmConsultCred.GrdCredito.PrintFooter = "/r" & font6 & "/n/n/r" & font5
''        VwGrid.hWndSpread = FrmConsultCred.GrdCredito.hWnd
''
''    ElseIf VGStrImprimir = "maqcli" Then
''        FrmMaqCli.Enabled = False
''        font5 = "Total de " & FrmMaqCli.GrdMaq.PrintPageCount & " página(s)"
''        font6 = FormataData(Date) & "   " & Time
''        FrmMaqCli.GrdMaq.PrintFooter = "/r" & font6 & "/n/n/r" & font5
''        VwGrid.hWndSpread = FrmMaqCli.GrdMaq.hWnd
''
''    ElseIf VGStrImprimir = "prod" Then
''        FrmPreco.Enabled = False
''        font5 = "Total de " & FrmPreco.GrdPreco.PrintPageCount & " página(s)"
''        font6 = FormataData(Date) & "   " & Time
''        FrmPreco.GrdPreco.PrintFooter = "/r" & font6 & "/n/n/r" & font5
''        VwGrid.hWndSpread = FrmPreco.GrdPreco.hWnd
''
''    End If
    
End Sub

